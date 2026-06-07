"""
ココナラ「ツール」検索結果スクレイパー (DrissionPage 版)
- 検索結果ページから個別サービスURLを収集（ページネーション）
- 各個別ページでタイトルと「販売実績数」を取得
- 販売実績数が MIN_SALES を超えるもののみ Excel に出力

[DrissionPage 使用]
- 実 Chromium / Edge ブラウザで stealth 自動化
- サーバーバーデン軽減のため、間隔ジッタ (4±2秒) + no_imgs
- ページネーション/個別ページ共にステートレスに逐次 + 並列

[リトライ方針: 全廃]
- 同じ IP で何度も叩くと Cloudflare のブロックが強化されるため、リトライは一切しない
- 1回だけ試して 403/Forbidden を検知したら諦め、ログに残す
- ネットワークエラー (timeout 等) も即諦め

[再開機能]
- 取得結果は cache.json に逐次保存
- 再実行時は未取得分のみ取得し、Excel を再生成
- コマンドライン引数でフェーズ制御:
    --collect       : URL収集のみ（フェーズ1）
    --fetch         : 個別ページ取得（フェーズ2、並列）
    --fetch-serial  : 個別ページ取得（フェーズ2、逐次）
    --excel         : Excel生成のみ（フェーズ3）
    引数なし        : 全フェーズを順次実行（再開）

[動作環境]
- Linux/Windows/macOS。GitHub Actions (ubuntu-latest) でも動作。
- システムに Chromium / Edge が必要（CI では apt-get install chromium-browser）
- Windows は Edge を優先（Chrome は --remote-debugging-port 互換性問題で DrissionPage から起動不可）
"""
from __future__ import annotations

import argparse
import json
import multiprocessing as mp
import os
import random
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urlencode

from DrissionPage import ChromiumOptions, ChromiumPage
from DrissionPage.errors import ElementNotFoundError
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

# ===== 設定 =====
KEYWORD = "ツール"
BASE_URL = "https://coconala.com"
SEARCH_PATH = "/contents_market/search"
CONTENT_KIND = "all"            # 全カテゴリ対象
MIN_SALES = 3                   # 販売実績数が MIN_SALES 以下のものは除外
OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_FILE = OUTPUT_DIR / "coconala_tools.xlsx"
URLS_FILE = OUTPUT_DIR / "urls_cache.json"        # フェーズ1の URL リスト
DATA_FILE = OUTPUT_DIR / "data_cache.json"        # フェーズ2の {url: {title, sales_count}}（マージ済み）

# 並列フェッチ
NUM_WORKERS = 3                 # 子プロセス数（DrissionPage Chromium をそれぞれ起動）
WORKER_FILE_TPL = "data_cache_w{}.json"

# サーバーバーデン軽減
REQUEST_INTERVAL = 4.0          # 基本待機秒
JITTER = 2.0                    # ±ランダム秒
PAGE_LOAD_WAIT = 3.0            # ページ読み込み後、余分に待つ秒（JSレンダリング待ち）
TIMEOUT_S = 30                  # タイムアウト（秒、DrissionPage は秒単位）

# ページネーション
MAX_PAGES = 60                  # 1923件 / 40件/ページ ≒ 49ページ
CACHE_SAVE_EVERY = 10           # N件ごとにキャッシュ保存


# ===== DrissionPage オプション =====
def make_options() -> ChromiumOptions:
    """DrissionPage のオプション設定（stealth 強化）

    Linux + CI (GitHub Actions) では --no-sandbox 必須。
    Windows では Edge / Chrome のパスを自動検出（環境変数 BROWSER_PATH で上書き可）。
    """
    co = ChromiumOptions()

    # ヘッドレス: 新方式 (--headless=new) は検知されにくい
    co.headless(True)
    co.set_argument("--headless=new")
    co.set_argument("--disable-blink-features=AutomationControlled")
    co.set_argument("--no-sandbox")
    co.set_argument("--disable-dev-shm-usage")
    co.set_argument("--disable-gpu")
    co.set_argument("--lang=ja-JP")

    # 画像読み込み無効化（高速化 + サーバーバーデン軽減）
    co.no_imgs()

    # User-Agent を明示（2024年の Chrome 安定版に偽装）
    ua = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/131.0.0.0 Safari/537.36"
    )
    co.set_user_agent(ua)

    # ブラウザパス: 環境変数 BROWSER_PATH があれば優先、なければ OS ごとに既定値
    # 注: Windows の Chrome は --remote-debugging-port と相性が悪い（DrissionPage で起動失敗）ため、
    #     Windows では Edge を優先する。
    browser_path = os.environ.get("BROWSER_PATH")
    if not browser_path:
        if sys.platform.startswith("win"):
            candidates = [
                r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            ]
            for c in candidates:
                if Path(c).exists():
                    browser_path = c
                    break
        elif sys.platform == "darwin":
            browser_path = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
        else:
            # Linux (GitHub Actions): chromium-browser
            for c in ("/usr/bin/chromium-browser", "/usr/bin/chromium", "/usr/bin/google-chrome"):
                if Path(c).exists():
                    browser_path = c
                    break
    if browser_path:
        co.set_browser_path(browser_path)
        print(f"  [INFO] ブラウザ: {browser_path}")

    return co


# ===== データ構造 =====
@dataclass
class Service:
    title: str
    sales_count: int
    url: str


# ===== URL/データ I/O =====
def build_search_url(keyword: str, page: int) -> str:
    params = {"content_kind": CONTENT_KIND, "keyword": keyword}
    if page > 1:
        params["page"] = page
    return f"{BASE_URL}{SEARCH_PATH}?{urlencode(params)}"


def polite_sleep(base: float = REQUEST_INTERVAL) -> None:
    """サーバーバーデン軽減のため、ジッタ付きで待機"""
    time.sleep(max(0.5, base + random.uniform(-JITTER, JITTER)))


def load_urls() -> list[str]:
    if URLS_FILE.exists():
        return json.loads(URLS_FILE.read_text(encoding="utf-8"))
    return []


def save_urls(urls: list[str]) -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    URLS_FILE.write_text(json.dumps(urls, ensure_ascii=False, indent=2), encoding="utf-8")


def load_data() -> dict[str, dict]:
    if DATA_FILE.exists():
        return json.loads(DATA_FILE.read_text(encoding="utf-8"))
    return {}


def save_data(data: dict[str, dict]) -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    DATA_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


# ===== ページ操作（DrissionPage） =====
def collect_service_urls_from_page(page: ChromiumPage) -> list[str]:
    """検索結果ページからコンテンツマーケットの商品URLを抽出。"""
    urls: set[str] = set()

    raw_hrefs = page.run_js(
        """() => {
            const anchors = document.querySelectorAll(
                'a[href*="/contents_market/pictures/"], a[href*="/contents_market/articles/"]'
            );
            return Array.from(anchors).map(a => a.getAttribute('href'));
        }"""
    ) or []
    pat = re.compile(r"^/contents_market/(pictures|articles)/[\w-]+$")
    for href in raw_hrefs:
        if not href or not pat.match(href):
            continue
        if href.startswith("/"):
            href = BASE_URL + href
        href = href.split("#", 1)[0]
        urls.add(href)

    return sorted(urls)


def get_total_count_from_page(page: ChromiumPage) -> int | None:
    """検索結果ヘッダーの「N件中」表示から総件数を取得。取れなければ None。"""
    try:
        text = page.run_js("() => document.body.innerText || ''") or ""
    except Exception:
        return None
    m = re.search(r"([\d,]+)\s*件中", text)
    if m:
        try:
            return int(m.group(1).replace(",", ""))
        except ValueError:
            return None
    return None


def fetch_service(page: ChromiumPage, url: str) -> Service | None:
    """個別商品ページからタイトルと販売実績数を取得（コンテンツマーケット用）

    【リトライなしポリシー】
    同じ IP / ブラウザ指紋で何度も叩くと Cloudflare のブロックが強化されるため、
    1回だけ試行して 403/Forbidden を検知したら即 None を返す。
    """
    try:
        page.get(url, timeout=TIMEOUT_S)
    except Exception as e:
        print(f"  [WARN] goto失敗: {url} ({e})")
        return None

    # og:title が出るまで待つ（最大15秒）。403/Forbidden が出たら即諦め。
    deadline = time.time() + 15.0
    og_content = ""
    got_valid_title = False
    while time.time() < deadline:
        try:
            og = page.ele('css:meta[property="og:title"]')
            if og:
                content = og.attr("content") or ""
                if content:
                    if "403" in content or "Forbidden" in content:
                        print(f"  [403] {url}")
                        return None
                    og_content = content
                    got_valid_title = True
                    break
        except Exception:
            pass
        time.sleep(0.5)

    if not got_valid_title:
        print(f"  [WARN] og:title未取得: {url}")
        return None

    # タイトル（og:title 優先、空なら page.title）
    title = og_content or (page.title or "")
    title = re.sub(r"\s*\|\s*ココナラコンテンツマーケット\s*$", "", title).strip()
    title = re.sub(r"\s*\|\s*ココナラ\s*$", "", title).strip()

    # page.title が 403 の場合の保険
    if "403" in title or "Forbidden" in title:
        print(f"  [403] {url}")
        return None

    # 「販売実績」テキストが出るまで追加で待つ（最大15秒）
    try:
        page.wait.ele_displayed("text=販売実績", timeout=15)
    except Exception:
        pass
    time.sleep(0.5)

    # 販売実績数: ページ本文から「販売実績 N件」(この商品の販売数) を抽出。
    # ※「販売実績 0」(出品者プロファイル側の累計) は "件" が付かないので区別できる。
    sales_count = 0
    try:
        text = page.run_js("() => document.body.innerText || ''") or ""
        m = re.search(r"販売.{0,5}?([\d,]+)\s*件", text)
        if not m:
            m = re.search(r"販売.{0,5}?([\d,]+)", text)
        if not m:
            m = re.search(r"売上数\s*([\d,]+)", text)
        if m:
            sales_count = int(m.group(1).replace(",", ""))
    except Exception:
        pass

    return Service(title=title, sales_count=sales_count, url=url)


# ===== Excel 出力 =====
def save_to_excel(services: list[Service], path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "ココナラ_ツール"

    headers = ["No.", "タイトル", "販売実績数", "URL"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for i, s in enumerate(services, 1):
        ws.cell(row=i + 1, column=1, value=i)
        ws.cell(row=i + 1, column=2, value=s.title)
        ws.cell(row=i + 1, column=3, value=s.sales_count)
        url_cell = ws.cell(row=i + 1, column=4, value=s.url)
        url_cell.hyperlink = s.url
        url_cell.font = Font(color="0563C1", underline="single")

    # 列幅
    widths = {"A": 6, "B": 65, "C": 14, "D": 55}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # ヘッダー行を固定
    ws.freeze_panes = "A2"

    path.parent.mkdir(exist_ok=True)
    wb.save(path)


# ===== フェーズ =====
def phase_collect(page: ChromiumPage) -> list[str]:
    """フェーズ1: 検索結果から個別商品URLを収集（リトライなし）"""
    all_urls = load_urls()
    print(f"[COLLECT] 既存キャッシュ: {len(all_urls)} 件")

    total_reported: int | None = None  # ページヘッダーから取得した総件数
    empty_streak = 0  # 連続0件ページ数

    for page_num in range(1, MAX_PAGES + 1):
        url = build_search_url(KEYWORD, page_num)
        print(f"\n[検索] ページ {page_num}: {url}")
        try:
            page.get(url, timeout=TIMEOUT_S)
        except Exception as e:
            print(f"  [WARN] goto失敗 ({e})。次のページへ。")
            polite_sleep()
            continue

        # 商品カードがDOMに現れるまで待つ（最大45秒、リトライなし）
        try:
            page.wait.ele_displayed(
                'css:a[href*="/contents_market/pictures/"], css:a[href*="/contents_market/articles/"]',
                timeout=45,
            )
        except Exception:
            print(f"  [WARN] 商品カードが見つかりません（JS読み込み遅延）次のページへ。")
            polite_sleep()
            continue
        time.sleep(1.0)

        # ページ1で総件数を取得
        if page_num == 1:
            total_reported = get_total_count_from_page(page)
            if total_reported:
                print(f"  [INFO] 検索総件数: {total_reported} 件")

        page_urls = collect_service_urls_from_page(page)
        if not page_urls:
            print(f"  → サービスが見つかりません。次のページへ。")
            polite_sleep()
            continue

        existing = set(all_urls)
        new_urls = [u for u in page_urls if u not in existing]
        all_urls.extend(new_urls)
        print(f"  → {len(new_urls)} 件取得（累計: {len(all_urls)}）")

        if len(new_urls) == 0:
            empty_streak += 1
            if empty_streak >= 2:
                print(f"  → 新規URLが{empty_streak}回連続で0件のため最終ページと判断")
                break
        else:
            empty_streak = 0

        polite_sleep()

    save_urls(all_urls)
    msg = f"\n[収集完了] 合計 {len(all_urls)} 件のサービスURL"
    if total_reported:
        msg += f"（ココナラ表示: {total_reported} 件）"
    msg += f" → {URLS_FILE.name}"
    print(msg)
    return all_urls


def phase_fetch_serial(page: ChromiumPage, all_urls: list[str]) -> dict[str, dict]:
    """フェーズ2: 個別ページからタイトル・販売実績を取得（未取得分のみ、逐次）"""
    data = load_data()
    to_fetch = [u for u in all_urls if u not in data]
    total = len(all_urls)
    done_before = total - len(to_fetch)
    print(f"[FETCH] 取得済み: {done_before} 件 / 未取得: {len(to_fetch)} 件 / 全体: {total} 件")

    for i, url in enumerate(to_fetch, 1):
        overall = done_before + i
        print(f"\n[取得] {overall}/{total}: {url}")
        service = fetch_service(page, url)
        if service:
            data[url] = {"title": service.title, "sales_count": service.sales_count}
            print(f"  タイトル: {service.title[:50]}")
            print(f"  販売実績: {service.sales_count} 件")
        else:
            data[url] = {"title": "(取得失敗)", "sales_count": 0}
        if i % CACHE_SAVE_EVERY == 0:
            save_data(data)
            print(f"  [cache saved] {len(data)} 件")
        if i < len(to_fetch):
            polite_sleep()

    save_data(data)
    print(f"\n[取得完了] {len(data)} 件保存 → {DATA_FILE.name}")
    return data


def _worker_main(worker_id: int, urls_chunk: list[str], output_dir: str) -> None:
    """子プロセス: 自分のチャンクのURLを取得して data_cache_w{worker_id}.json に保存"""
    out_path = Path(output_dir) / WORKER_FILE_TPL.format(worker_id)
    data: dict[str, dict] = {}
    if out_path.exists():
        try:
            data = json.loads(out_path.read_text(encoding="utf-8"))
        except Exception:
            data = {}
    already = sum(1 for u in urls_chunk if u in data)
    print(f"[w{worker_id}] 起動: 担当 {len(urls_chunk)} 件 / 既存 {already} 件", flush=True)

    page = ChromiumPage(make_options())
    page.set.timeouts(base=TIMEOUT_S, page_load=TIMEOUT_S)

    try:
        for i, url in enumerate(urls_chunk, 1):
            if url in data:
                continue
            try:
                service = fetch_service(page, url)
                if service:
                    data[url] = {"title": service.title, "sales_count": service.sales_count}
                    print(f"  [w{worker_id}] {i}/{len(urls_chunk)} OK {service.sales_count}件: {service.title[:40]}", flush=True)
                else:
                    data[url] = {"title": "(取得失敗)", "sales_count": 0}
                    print(f"  [w{worker_id}] {i}/{len(urls_chunk)} FAIL: {url}", flush=True)
            except Exception as e:
                print(f"  [w{worker_id}] {i}/{len(urls_chunk)} ERROR: {url} ({e})", flush=True)
                data[url] = {"title": "(取得失敗)", "sales_count": 0}

            if i % CACHE_SAVE_EVERY == 0:
                out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

            time.sleep(max(0.5, REQUEST_INTERVAL + random.uniform(-JITTER, JITTER)))
    finally:
        try:
            page.quit()
        except Exception:
            pass

    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[w{worker_id}] 完了: {len(data)} 件保存", flush=True)


def phase_fetch_parallel(all_urls: list[str], num_workers: int = NUM_WORKERS) -> dict[str, dict]:
    """フェーズ2（並列）: NUM_WORKERS 個の子プロセスで並列取得"""
    data = load_data()
    print(f"[FETCH] 既存: {len(data)} 件 / 全体: {len(all_urls)} 件 / ワーカー数: {num_workers}")

    # ラウンドロビンで安定分割（再実行時も同じURLが同じワーカーへ）
    chunks: list[list[str]] = [[] for _ in range(num_workers)]
    for i, url in enumerate(all_urls):
        chunks[i % num_workers].append(url)
    for i, c in enumerate(chunks):
        print(f"  worker{i}: 担当 {len(c)} 件")

    processes = []
    for i in range(num_workers):
        p = mp.Process(
            target=_worker_main,
            args=(i, chunks[i], str(OUTPUT_DIR)),
        )
        p.start()
        processes.append(p)

    for p in processes:
        p.join()

    # マージ
    for i in range(num_workers):
        wf = OUTPUT_DIR / WORKER_FILE_TPL.format(i)
        if wf.exists():
            try:
                d = json.loads(wf.read_text(encoding="utf-8"))
                data.update(d)
            except Exception as e:
                print(f"  [WARN] {wf.name} 読み込み失敗: {e}")

    save_data(data)
    print(f"\n[取得完了] {len(data)} 件マージ → {DATA_FILE.name}")
    return data


def phase_excel(data: dict[str, dict]) -> list[Service]:
    """フェーズ3: フィルタリングとExcel生成"""
    services = [
        Service(title=v["title"], sales_count=v["sales_count"], url=u)
        for u, v in data.items()
    ]
    filtered = [s for s in services if s.sales_count > MIN_SALES]
    print(f"[EXCEL] 全 {len(services)} 件 / MIN_SALES={MIN_SALES} 超 {len(filtered)} 件")

    filtered.sort(key=lambda s: s.sales_count, reverse=True)
    save_to_excel(filtered, OUTPUT_FILE)
    print(f"[完了] Excel出力: {OUTPUT_FILE}")
    return filtered


# ===== メイン =====
def parse_args(argv: list[str]) -> argparse.Namespace:
    p = argparse.ArgumentParser()
    p.add_argument("--collect", action="store_true", help="URL収集のみ")
    p.add_argument("--fetch", action="store_true", help="個別ページ取得のみ（並列）")
    p.add_argument("--fetch-serial", action="store_true", help="個別ページ取得のみ（逐次）")
    p.add_argument("--excel", action="store_true", help="Excel生成のみ")
    p.add_argument("--test", type=int, default=0, metavar="N", help="テストモード：最初のN件のみ")
    p.add_argument("--workers", type=int, default=NUM_WORKERS, help=f"並列ワーカー数 (default: {NUM_WORKERS})")
    return p.parse_args(argv[1:])


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv if argv is None else argv)
    test_limit = args.test if args.test > 0 else None

    print(f"[START] keyword='{KEYWORD}', MIN_SALES={MIN_SALES}, workers={args.workers}")
    if test_limit:
        print(f"[TEST MODE] 最初の {test_limit} 件のみ")

    only = [args.collect, args.fetch, args.fetch_serial, args.excel].count(True)
    single_mode = only > 0

    all_urls: list[str] = []
    data: dict[str, dict] = {}

    needs_browser = (not single_mode) or args.collect or args.fetch_serial

    if needs_browser:
        page = ChromiumPage(make_options())
        page.set.timeouts(base=TIMEOUT_S, page_load=TIMEOUT_S)
        try:
            if not single_mode or args.collect:
                all_urls = phase_collect(page)
                if test_limit:
                    all_urls = all_urls[:test_limit]

            if args.fetch_serial:
                if not all_urls:
                    all_urls = load_urls()
                    if test_limit:
                        all_urls = all_urls[:test_limit]
                data = phase_fetch_serial(page, all_urls)
        finally:
            try:
                page.quit()
            except Exception:
                pass

    if not single_mode or args.fetch:
        if not all_urls:
            all_urls = load_urls()
            if test_limit:
                all_urls = all_urls[:test_limit]
        data = phase_fetch_parallel(all_urls, num_workers=args.workers)

    if not single_mode or args.excel:
        if not data:
            data = load_data()
        phase_excel(data)

    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
