from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import logging
import base64
from datetime import datetime
import argparse
import json
import re

# ロギングの設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("receipt_download_manual.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# グローバル変数としてdownload_dirを定義
download_dir = None

# ダウンロードディレクトリの設定を修正
def create_download_dir():
    """タイムスタンプ付きのダウンロードディレクトリを作成する"""
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    download_dir = os.path.join(os.getcwd(), f'receipts_{timestamp}')
    os.makedirs(download_dir, exist_ok=True)
    logger.info(f"ダウンロードディレクトリを作成しました: {download_dir}")
    return download_dir

# PDFファイル名の生成関数を修正
def generate_pdf_filename(index, receipt_number=None):
    """PDFファイル名を生成する（通し番号と領収書/請求書番号を組み合わせる）"""
    # 日付を取得
    date_str = datetime.now().strftime("%Y%m%d")
    
    # 通し番号部分（必ず含める）
    index_str = f"{index:03d}"
    
    # 領収書/請求書番号がある場合はそれを組み合わせる
    if receipt_number:
        # 番号が長すぎる場合は短くする
        if len(receipt_number) > 20:
            receipt_number = receipt_number[:20]
        return f"領収書_{index_str}_{receipt_number}.pdf"
    else:
        # 番号がない場合は通し番号のみ
        return f"領収書_{index_str}.pdf"

# ダウンロードの完了を待機する関数
def wait_for_download_complete(directory, timeout=30):
    """ダウンロードの完了を待機する関数"""
    start_time = time.time()
    initial_files = set(os.listdir(directory))
    
    while time.time() - start_time < timeout:
        current_files = set(os.listdir(directory))
        new_files = current_files - initial_files
        
        # 新しいファイルが見つかった場合
        if new_files:
            # .crdownload や .tmp ファイルがなければダウンロード完了
            if not any(f.endswith('.crdownload') or f.endswith('.tmp') for f in new_files):
                return list(new_files)
        
        time.sleep(1)
    
    return []

# 一覧ページに戻る関数を修正
def go_back_to_list_page(driver, page_num=1):
    """一覧ページに確実に戻るための関数"""
    max_attempts = 3
    for attempt in range(max_attempts):
        try:
            logger.info(f"支払一覧ページ {page_num} に移動中... ({attempt+1}/{max_attempts})")
            
            # 直接URLで移動（最も確実な方法）
            if page_num == 1:
                driver.get('https://crowdworks.jp/payments?ref=login_header')
            else:
                driver.get(f'https://crowdworks.jp/payments?page={page_num}&ref=login_header')
            
            # ページの読み込みを待機（複数の検索方法を試す）
            try:
                # 方法1: テキストで検索
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//a[contains(text(), '領収書')]"))
                )
                logger.info(f"支払一覧ページ {page_num} に移動完了（領収書リンク検出）")
                return True
            except:
                logger.info("テキストでの領収書リンク検索に失敗しました")
            
            try:
                # 方法2: クラス名で検索
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "a.text-button.issuable"))
                )
                logger.info(f"支払一覧ページ {page_num} に移動完了（クラス名で検出）")
                return True
            except:
                logger.info("クラス名での領収書リンク検索に失敗しました")
            
            try:
                # 方法3: href属性で検索
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//a[contains(@href, '/receipt_sheets/new')]"))
                )
                logger.info(f"支払一覧ページ {page_num} に移動完了（href属性で検出）")
                return True
            except:
                logger.info("href属性での領収書リンク検索に失敗しました")
            
            # ページのタイトルやURLで確認
            if 'payments' in driver.current_url:
                logger.info(f"支払一覧ページ {page_num} に移動完了（URL確認）")
                return True
            
            logger.warning(f"支払一覧ページ {page_num} の確認に失敗しました (試行 {attempt+1}/{max_attempts})")
            time.sleep(3)
            
        except Exception as e:
            logger.warning(f"支払一覧ページ {page_num} への移動に失敗 (試行 {attempt+1}/{max_attempts}): {str(e)}")
            time.sleep(3)
    
    logger.error(f"支払一覧ページ {page_num} への移動が最大試行回数を超えました")
    return False

def download_receipts_with_manual_login(download_dir=None, config=None):
    """手動ログインを組み込んだ領収書ダウンロード処理"""
    # Chromeの設定と初期化
    driver = setup_chrome_driver()
    
    try:
        # ログイン処理
        if not perform_manual_login(driver):
            print("ログインに失敗しました。処理を終了します。")
            return
        
        # 一覧ページのURLを入力してもらう
        print("\n=== 領収書一覧ページの設定 ===")
        print("1: デフォルトの一覧ページを使用する (https://crowdworks.jp/payments)")
        print("2: カスタムURLを入力する")
        choice = input("選択してください (1/2): ")
        
        if choice == "2":
            list_page_url = input("領収書一覧ページのURLを入力してください: ")
        else:
            list_page_url = "https://crowdworks.jp/payments?ref=login_header"
        
        # 一覧ページに移動
        logger.info(f"一覧ページ {list_page_url} にアクセスします")
        driver.get(list_page_url)
        wait_for_page_load(driver)
        
        # 総ページ数と総領収書数を取得
        print("\n領収書の総数を計算中...")
        total_pages, total_receipts = get_total_pages_and_receipts(driver)
        if total_pages > 0:
            logger.info(f"合計 {total_pages} ページ、{total_receipts} 件の領収書が見つかりました")
            print(f"\n合計 {total_pages} ページ、{total_receipts} 件の領収書が見つかりました")
        else:
            logger.info("ページ数を自動検出できませんでした。すべてのページを処理します。")
        
        # すべてのページを処理
        process_all_pages(driver, total_pages, total_receipts)
        
    except Exception as e:
        logger.error(f"処理中にエラーが発生しました: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        print("\n処理中にエラーが発生しました。詳細はログファイルを確認してください。")
    finally:
        # ブラウザを閉じる
        driver.quit()
        logger.info("ブラウザを閉じました")

def setup_chrome_driver():
    """ChromeDriverの設定と初期化を行う"""
    global download_dir
    
    # Chromeの設定
    options = Options()
    options.add_argument("--window-size=1920,1080")
    
    # ダウンロード設定
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    
    # WebDriverの初期化
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)

def perform_manual_login(driver):
    """手動ログイン処理を行う"""
    try:
        # ログインページにアクセス
        logger.info("ログインページにアクセスしています...")
        driver.get('https://crowdworks.jp/login')
        
        # 手動ログインを待機
        logger.info("手動でログインしてください。マイページが表示されるまで待機します...")
        print("=== 手動でログインしてください ===")
        print("ログインが完了したら、自動的に処理が続行されます")
        
        # マイページが表示されるまで待機
        WebDriverWait(driver, 300).until(  # 5分間待機
            lambda d: 'マイページ' in d.page_source or '/mypage' in d.current_url
        )
        
        logger.info("ログインを確認しました。処理を続行します")
        return True
    except Exception as e:
        logger.error(f"ログイン処理中にエラーが発生: {str(e)}")
        return False

def get_total_pages_and_receipts(driver):
    """総ページ数と総領収書数を取得する"""
    try:
        # 方法1: ページネーションの最後のページ番号を取得
        pagination_links = driver.find_elements(By.CSS_SELECTOR, ".pagination a, .pager a")
        page_numbers = []
        
        for link in pagination_links:
            try:
                text = link.text.strip()
                if text.isdigit():
                    page_numbers.append(int(text))
            except:
                pass
        
        total_pages = max(page_numbers) if page_numbers else 1
        
        # 現在のページの領収書数を取得
        current_page_receipts = len(get_receipt_links(driver))
        
        # 全ページの領収書数を推定
        # 最後のページ以外は同じ数の領収書があると仮定
        if total_pages > 1:
            # 最後のページの領収書数を取得するために一時的に移動
            last_page_url = f"https://crowdworks.jp/payments?page={total_pages}&ref=login_header"
            current_url = driver.current_url
            
            driver.get(last_page_url)
            wait_for_page_load(driver)
            last_page_receipts = len(get_receipt_links(driver))
            
            # 元のページに戻る
            driver.get(current_url)
            wait_for_page_load(driver)
            
            # 総領収書数を計算
            total_receipts = current_page_receipts * (total_pages - 1) + last_page_receipts
        else:
            total_receipts = current_page_receipts
        
        logger.info(f"ページネーションから {total_pages} ページ、合計 {total_receipts} 件の領収書を検出しました")
        return total_pages, total_receipts
        
    except Exception as e:
        logger.error(f"ページ数と領収書数の取得に失敗しました: {str(e)}")
        
        # 方法3: ユーザーに尋ねる
        print("\n=== ページ数と領収書数の設定 ===")
        print("自動的にページ数と領収書数を検出できませんでした。")
        try:
            user_pages = int(input("処理する総ページ数を入力してください (不明な場合は1): ") or "1")
            user_receipts = int(input("処理する総領収書数を入力してください (不明な場合は推定します): ") or "0")
            
            if user_receipts <= 0:
                # 現在のページの領収書数から推定
                current_page_receipts = len(get_receipt_links(driver))
                user_receipts = current_page_receipts * user_pages
            
            return user_pages, user_receipts
        except:
            # デフォルト値
            return 1, len(get_receipt_links(driver))

def collect_page_urls(driver):
    """全ページのURLを収集する"""
    page_urls = []
    current_url = driver.current_url
    page_urls.append(current_url)
    logger.info(f"1ページ目のURLを保存: {current_url}")
    
    page_num = 1
    while True:
        try:
            # 「次へ」リンクを探す
            next_link = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((
                    By.XPATH, 
                    "//a[contains(text(), '次へ')] | //a[contains(text(), '次の')] | //a[@rel='next']"
                ))
            )
            
            # リンクが表示されるまでスクロール
            driver.execute_script("arguments[0].scrollIntoView(true);", next_link)
            time.sleep(1)
            
            # 「次へ」リンクをクリック
            if safe_click(driver, next_link):
                wait_for_page_load(driver)
                time.sleep(2)
                
                # 新しいページのURLを保存
                page_num += 1
                current_url = driver.current_url
                page_urls.append(current_url)
                logger.info(f"{page_num}ページ目のURLを保存: {current_url}")
                
                # 領収書リンクの存在を確認
                receipt_links = get_receipt_links(driver)
                if not receipt_links:
                    logger.info(f"ページ {page_num} に領収書が見つかりません。URL収集を終了します。")
                    break
            else:
                logger.info("「次へ」リンクのクリックに失敗しました。URL収集を終了します。")
                break
                
        except Exception as e:
            logger.info(f"次のページが見つかりません。URL収集を終了します: {str(e)}")
            break
    
    logger.info(f"合計 {len(page_urls)} ページのURLを収集しました")
    return page_urls

def process_all_pages(driver, total_pages=0, total_receipts=0):
    """すべてのページの領収書を処理する"""
    # まず全ページのURLを収集
    logger.info("全ページのURLを収集します...")
    page_urls = collect_page_urls(driver)
    total_pages = len(page_urls)
    logger.info(f"収集したページ数: {total_pages}")
    
    page_num = 1
    total_downloaded = 0
    current_receipt_index = 0
    max_retries = 3
    
    # 収集したURLを使って各ページを処理
    for page_url in page_urls:
        logger.info(f"ページ {page_num}/{total_pages} の処理を開始します")
        retry_count = 0
        
        while retry_count < max_retries:
            try:
                # URLを使って直接ページに移動
                driver.get(page_url)
                wait_for_page_load(driver)
                time.sleep(2)
                
                # ページ内の領収書を取得
                receipt_links = get_receipt_links(driver)
                if not receipt_links:
                    logger.error(f"ページ {page_num} で領収書リンクが見つかりません（試行 {retry_count + 1}/{max_retries}）")
                    retry_count += 1
                    time.sleep(3)
                    continue
                
                page_receipts = len(receipt_links)
                logger.info(f"ページ {page_num} で {page_receipts} 件の領収書を検出しました")
                
                # ページ内の領収書を処理
                for i in range(1, page_receipts + 1):
                    current_receipt_index += 1
                    
                    # 進捗表示
                    if total_receipts > 0:
                        display_progress(current_receipt_index, total_receipts, "領収書ダウンロード")
                    else:
                        display_progress(i, page_receipts, f"ページ {page_num}/{total_pages} の領収書ダウンロード")
                    
                    # 領収書の処理
                    receipt_retry_count = 0
                    while receipt_retry_count < max_retries:
                        try:
                            if i > 1:
                                # 一覧ページに戻る
                                driver.get(page_url)
                                wait_for_page_load(driver)
                                time.sleep(2)
                            
                            success = process_receipt_by_index(driver, i, page_receipts, current_receipt_index)
                            if success:
                                total_downloaded += 1
                                break
                            else:
                                receipt_retry_count += 1
                                time.sleep(3)
                                
                        except Exception as e:
                            logger.error(f"領収書 {i} の処理中にエラー: {str(e)}")
                            receipt_retry_count += 1
                            if receipt_retry_count >= max_retries:
                                if not handle_receipt_error(driver, e, i, page_num):
                                    return total_downloaded
                            time.sleep(3)
                
                # ページの処理が成功したらループを抜ける
                break
                
            except Exception as e:
                logger.error(f"ページ {page_num} の処理中にエラー: {str(e)}")
                retry_count += 1
                if retry_count >= max_retries:
                    print(f"\nページ {page_num} の処理中にエラーが発生しました。")
                    print("1: 再試行する")
                    print("2: このページをスキップして次に進む")
                    print("3: 処理を中止する")
                    choice = input("選択してください (1/2/3): ")
                    
                    if choice == "1":
                        retry_count = 0
                        continue
                    elif choice == "3":
                        return total_downloaded
                    else:
                        break
                time.sleep(3)
        
        page_num += 1
    
    # 最終結果を表示
    if total_receipts > 0:
        success_rate = (total_downloaded / total_receipts) * 100
        print(f"\n処理完了: 合計 {total_downloaded}/{total_receipts} 件の領収書をダウンロードしました ({success_rate:.1f}%)")
    else:
        print(f"\n処理完了: 合計 {total_downloaded} 件の領収書をダウンロードしました")
    
    logger.info(f"処理完了: 合計 {total_downloaded} 件の領収書をダウンロードしました")
    return total_downloaded

def go_to_page(driver, page_num):
    """指定したページ番号に直接移動する"""
    try:
        url = f"https://crowdworks.jp/payments?page={page_num}&ref=login_header"
        logger.info(f"ページ {page_num} に直接移動します: {url}")
        
        driver.get(url)
        wait_for_page_load(driver)
        
        # 領収書リンクがあるか確認
        try:
            receipt_links = get_receipt_links(driver)
            if receipt_links:
                logger.info(f"ページ {page_num} に {len(receipt_links)} 件の領収書が見つかりました")
                return True
            else:
                logger.info(f"ページ {page_num} に領収書が見つかりませんでした")
                return False
        except Exception as e:
            logger.error(f"ページ {page_num} での領収書リンク検索に失敗しました: {str(e)}")
            return False
    except Exception as e:
        logger.error(f"ページ {page_num} への直接移動に失敗しました: {str(e)}")
        return False

def navigate_to_page(driver, target_page):
    """「次へ」リンクを使って指定したページまで移動する"""
    current_page = 1
    
    while current_page < target_page:
        logger.info(f"ページ {current_page} から {target_page} へ移動中...")
        
        if not move_to_next_page(driver):
            logger.error(f"ページ {current_page + 1} への移動に失敗しました")
            return False
        
        current_page += 1
    
    return True

def process_page_receipts(driver, page_num):
    """ページ内の領収書を処理する"""
    try:
        # 領収書リンクの数を取得
        receipt_links = get_receipt_links(driver)
        
        if not receipt_links:
            # デバッグ用にページのHTMLを保存
            with open(os.path.join(download_dir, f"page_source_page_{page_num}.html"), "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            
            logger.info("このページには領収書がありません。処理を終了します。")
            return 0
        
        total_receipts = len(receipt_links)
        logger.info(f"ページ {page_num} で {total_receipts}件の領収書が見つかりました")
        
        # 既存のダウンロードファイル名を取得
        existing_files = set(os.listdir(download_dir))
        page_downloaded = 0
        current_receipt_index = 0  # current_receipt_indexを定義
        
        # 各領収書をダウンロード
        for i in range(1, total_receipts + 1):
            try:
                # 毎回一覧ページに戻ってから処理する
                if i > 1:  # 最初の領収書の場合は既に一覧ページにいるので不要
                    if not go_back_to_list_page(driver, page_num):
                        logger.error(f"領収書 {i} の処理前に一覧ページへの移動に失敗しました")
                        if not handle_navigation_error(driver, i, page_num):
                            return -1
                        continue
                
                current_receipt_index += 1  # インデックスをインクリメント
                # 領収書リンクを毎回新しく取得
                success = process_receipt_by_index(driver, i, total_receipts, current_receipt_index)
                if success:
                    page_downloaded += 1
                
            except Exception as e:
                if not handle_receipt_error(driver, e, i, page_num):
                    return -1  # 処理中止
        
        return page_downloaded
        
    except Exception as e:
        logger.error(f"ページ {page_num} の処理中にエラーが発生: {str(e)}")
        print(f"ページ {page_num} の処理中にエラーが発生しました。続行しますか？ (y/n)")
        if input().lower() != 'y':
            logger.info("ユーザーにより処理が中止されました")
            return -1
        return 0

def process_receipt_by_index(driver, index, total, global_index=None):
    """インデックスを指定して領収書を処理する"""
    # 通し番号を使用（指定されていない場合はページ内のインデックスを使用）
    actual_index = global_index if global_index is not None else index
    
    logger.info(f"領収書 {index}/{total} (通し番号: {actual_index}) の処理を開始します")
    
    try:
        # 領収書リンクを新しく取得
        current_links = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, "//a[contains(text(), '領収書')]"))
        )
        
        if index <= len(current_links):
            # index番目の領収書リンクを取得
            link = current_links[index-1]
            
            # リンクが表示されるまでスクロール
            driver.execute_script("arguments[0].scrollIntoView(true);", link)
            time.sleep(1)
            
            # リンクのURLを取得して直接アクセス（より安定した方法）
            try:
                href = link.get_attribute("href")
                if href:
                    logger.info(f"領収書 {index} (通し番号: {actual_index}) のURLに直接アクセスします: {href}")
                    driver.get(href)
                    wait_for_page_load(driver)
                    time.sleep(2)
                    
                    # ヘッダーを非表示にするJavaScriptを実行
                    driver.execute_script("""
                        // ヘッダーメッセージを非表示にする
                        var headers = document.querySelectorAll('.alert, .alert-success, .notice, .message, .flash-message');
                        headers.forEach(function(header) {
                            header.style.display = 'none';
                        });
                        
                        // 印刷用のスタイルを追加
                        var style = document.createElement('style');
                        style.innerHTML = `
                            @media print {
                                .alert, .alert-success, .notice, .message, .flash-message { display: none !important; }
                                body { margin: 0; padding: 0; }
                                * { -webkit-print-color-adjust: exact !important; }
                            }
                        `;
                        document.head.appendChild(style);
                    """)
                    time.sleep(1)
                    
                else:
                    # URLが取得できない場合はクリック
                    logger.info(f"領収書 {index} (通し番号: {actual_index}) のリンクをクリックします")
                    if not safe_click(driver, link):
                        logger.error(f"領収書リンク {index} (通し番号: {actual_index}) のクリックに失敗しました")
                        return False
                    wait_for_page_load(driver)
                    time.sleep(2)
                    
                    # ヘッダーを非表示にするJavaScriptを実行
                    driver.execute_script("""
                        // ヘッダーメッセージを非表示にする
                        var headers = document.querySelectorAll('.alert, .alert-success, .notice, .message, .flash-message');
                        headers.forEach(function(header) {
                            header.style.display = 'none';
                        });
                        
                        // 印刷用のスタイルを追加
                        var style = document.createElement('style');
                        style.innerHTML = `
                            @media print {
                                .alert, .alert-success, .notice, .message, .flash-message { display: none !important; }
                                body { margin: 0; padding: 0; }
                                * { -webkit-print-color-adjust: exact !important; }
                            }
                        `;
                        document.head.appendChild(style);
                    """)
                    time.sleep(1)
                    
            except Exception as e:
                logger.error(f"領収書 {index} (通し番号: {actual_index}) のURL取得に失敗: {str(e)}")
                return False

            # 既に発行済みかどうかを確認（印刷ボタンが存在するか）
            print_button = None
            try:
                print_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, ".print_button, .cw-button_action.print_button"))
                )
                logger.info("印刷ボタンを見つけました。既に発行済みの領収書です。")
            except:
                logger.info("印刷ボタンが見つかりませんでした。未発行の領収書です。")
            
            # 既に発行済みの場合はPDF保存処理へ、そうでなければ発行処理へ
            if print_button:
                # PDFとして保存（印刷ボタンを使わない）
                pdf_file_name = save_as_pdf(driver, actual_index)
                if pdf_file_name:
                    logger.info(f"領収書 {index} (通し番号: {actual_index}) を保存しました: {pdf_file_name}")
                    return True
                else:
                    # PDF保存に失敗した場合は代替方法を試す
                    logger.info("PDFとして保存に失敗しました。代替方法を試みます...")
                    
                    # 代替方法1: 再度PDFとして保存を試みる
                    time.sleep(2)  # 少し待機してから再試行
                    pdf_file_name = save_as_pdf(driver, actual_index)
                    if pdf_file_name:
                        logger.info(f"再試行で領収書 {index} (通し番号: {actual_index}) を保存しました: {pdf_file_name}")
                        return True
                    
                    # 代替方法2: スクリーンショットとして保存
                    logger.info("スクリーンショットとして保存を試みます...")
                    screenshot_path = os.path.join(download_dir, f"領収書_{actual_index:02d}.png")
                    driver.save_screenshot(screenshot_path)
                    logger.info(f"領収書 {index} (通し番号: {actual_index}) をスクリーンショットとして保存しました: {screenshot_path}")
                    return True
            else:
                # 未発行の領収書の場合、発行ボタンを探して処理
                try:
                    # まず「プレビューで内容を確認する」ボタンを探す
                    preview_button = None
                    try:
                        preview_button = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, "//input[@value='プレビューで内容を確認する'] | //button[contains(text(), 'プレビューで内容を確認する')]"))
                        )
                        logger.info("「プレビューで内容を確認する」ボタンを見つけました。クリックします。")
                        
                        if safe_click(driver, preview_button):
                            wait_for_page_load(driver)
                            time.sleep(2)
                            # プレビュー画面でヘッダーを非表示に
                            hide_header_elements(driver)
                            logger.info("プレビュー画面に移動しました。")
                    except Exception as e:
                        logger.info(f"「プレビューで内容を確認する」ボタンが見つからないか、クリックに失敗しました: {str(e)}")
                        logger.info("プレビュー画面をスキップして発行処理を続行します。")
                    
                    # 次に「この内容で発行する」ボタンを探す
                    issue_button = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//input[@value='この内容で発行する'] | //button[contains(text(), 'この内容で発行する')]"))
                    )
                    logger.info("「この内容で発行する」ボタンを見つけました。クリックします。")
                    
                    if safe_click(driver, issue_button):
                        wait_for_page_load(driver)
                        time.sleep(2)
                        # 発行直後にヘッダーを非表示に
                        hide_header_elements(driver)
                        
                        # 確認ダイアログの「はい」ボタンを探して処理
                        try:
                            yes_button = WebDriverWait(driver, 5).until(
                                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'はい')] | //input[@value='はい'] | //a[contains(text(), 'はい')] | //div[contains(@class, 'dialog')]//button[1]"))
                            )
                            logger.info("確認ダイアログの「はい」ボタンを見つけました。クリックします。")
                            
                            if safe_click(driver, yes_button):
                                wait_for_page_load(driver)
                                time.sleep(3)
                                # 発行完了後にもヘッダーを非表示に
                                hide_header_elements(driver)
                                
                                # 発行後のページをPDFとして保存
                                pdf_file_name = save_as_pdf(driver, actual_index)
                                if pdf_file_name:
                                    logger.info(f"領収書 {index} を保存しました: {pdf_file_name}")
                                    return True
                                else:
                                    # PDF保存に失敗した場合はスクリーンショットを取る
                                    logger.info("PDFとして保存に失敗しました。スクリーンショットを取ります...")
                                    screenshot_path = os.path.join(download_dir, f"領収書_{actual_index:02d}.png")
                                    driver.save_screenshot(screenshot_path)
                                    logger.info(f"領収書 {index} をスクリーンショットとして保存しました: {screenshot_path}")
                                    return True
                        except Exception as e:
                            logger.error(f"確認ダイアログの「はい」ボタン処理中にエラー: {str(e)}")
                            # JavaScriptでの処理を試みる
                            try:
                                driver.execute_script("""
                                    // 確認ダイアログの「はい」ボタンを探して自動的にクリック
                                    var yesButtons = document.querySelectorAll('button, input[type="button"], input[type="submit"], a.button');
                                    for (var i = 0; i < yesButtons.length; i++) {
                                        var btn = yesButtons[i];
                                        if (btn.textContent.includes('はい') || btn.value === 'はい') {
                                            btn.click();
                                            return true;
                                        }
                                    }
                                    // 最初のボタンをクリック（多くの場合「はい」が最初）
                                    var firstButton = document.querySelector('.dialog button, .modal button, .confirm button');
                                    if (firstButton) {
                                        firstButton.click();
                                        return true;
                                    }
                                    return false;
                                """)
                                time.sleep(3)
                                # JavaScript実行後にもヘッダーを非表示に
                                hide_header_elements(driver)
                            except Exception as js_error:
                                logger.error(f"JavaScriptによる確認ダイアログ処理に失敗: {str(js_error)}")

                        # 発行完了後の最終確認としてヘッダーを非表示に
                        hide_header_elements(driver)
                        
                        # 発行後のページをPDFとして保存
                        pdf_file_name = save_as_pdf(driver, actual_index)
                        if pdf_file_name:
                            logger.info(f"領収書 {index} を保存しました: {pdf_file_name}")
                            return True
                        else:
                            # PDF保存に失敗した場合はスクリーンショットを取る
                            logger.info("PDFとして保存に失敗しました。スクリーンショットを取ります...")
                            screenshot_path = os.path.join(download_dir, f"領収書_{actual_index:02d}.png")
                            driver.save_screenshot(screenshot_path)
                            logger.info(f"領収書 {index} をスクリーンショットとして保存しました: {screenshot_path}")
                            return True
                except Exception as e:
                    logger.error(f"発行ボタンの処理中にエラー: {str(e)}")
            
            # 手動操作を求める
            print("\n=== 自動処理に失敗しました ===")
            print("手動で領収書を発行・保存してください。")
            input("操作が完了したら、Enterキーを押して続行してください...")
            return True
        else:
            logger.error(f"インデックス {index} が領収書リンクの数 {len(current_links)} を超えています")
            return False
    except Exception as e:
        logger.error(f"領収書 {index} (通し番号: {actual_index}) の処理中にエラー: {str(e)}")
        raise
    
    return False

def handle_navigation_error(driver, index, page_num):
    """ナビゲーションエラーを処理する"""
    print(f"=== 領収書 {index} の処理前に一覧ページへの移動に失敗しました ===")
    print("1: 再試行する")
    print("2: この領収書をスキップして次に進む")
    print("3: 処理を中止する")
    choice = input("選択してください (1/2/3): ")
    
    if choice == "1":
        # 再試行
        logger.info("ユーザーが再試行を選択しました")
        return True
    elif choice == "3":
        logger.info("ユーザーにより処理が中止されました")
        return False
    else:
        # スキップ
        logger.info("ユーザーがスキップを選択しました")
        return True

def get_receipt_links(driver):
    """ページ内の領収書リンクを取得する（請求書を除外）"""
    try:
        # 領収書リンクを探す（テキストが完全に「領収書」のみのリンク）
        receipt_links = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, "//a[text()='領収書']"))
        )
        
        # 念のため、「請求書」を含まないことを確認
        filtered_links = []
        for link in receipt_links:
            text = link.text.strip()
            if text == '領収書' and '請求書' not in text:
                filtered_links.append(link)
        
        logger.info(f"テキストで {len(filtered_links)} 件の領収書リンクを見つかりました（請求書を除外）")
        return filtered_links
    except Exception as e:
        # 代替方法：より緩いXPathで検索し、テキストで絞り込む
        try:
            all_links = driver.find_elements(By.XPATH, "//a[contains(text(), '領収書')]")
            filtered_links = []
            
            for link in all_links:
                text = link.text.strip()
                if text == '領収書' and '請求書' not in text:
                    filtered_links.append(link)
            
            logger.info(f"代替方法で {len(filtered_links)} 件の領収書リンクを見つかりました（請求書を除外）")
            return filtered_links
        except Exception as e2:
            logger.error(f"領収書リンクの取得に失敗: {str(e)} / {str(e2)}")
            return []

def hide_header_elements(driver):
    """ヘッダー要素を非表示にする共通処理（強化版）"""
    try:
        # より広範囲のヘッダー要素を非表示にするJavaScriptを実行
        driver.execute_script("""
            // すべての通知、アラート、メッセージ要素を非表示にする
            var elementsToHide = document.querySelectorAll(
                '.alert, .alert-success, .alert-info, .alert-warning, .alert-danger, ' +
                '.notice, .message, .flash-message, .header-message, .notification, ' +
                '.status-message, div[role="alert"], .toast, .banner, ' +
                '.info-message, .success-message, .warning-message, .error-message, ' +
                '.flash, .flash-notice, .flash-success, .flash-error, ' +
                '.message-container, .message-box, .notification-container, ' +
                '.close-button, .close-icon, button.close, ' +
                'div[class*="alert"], div[class*="notice"], div[class*="message"], ' +
                'div[class*="notification"], div[class*="toast"], ' +
                'div[class*="info"], div[class*="success"], div[class*="warning"], div[class*="error"]'
            );
            
            // すべての要素を非表示に
            elementsToHide.forEach(function(element) {
                element.style.display = 'none';
                element.style.visibility = 'hidden';
                element.style.opacity = '0';
                element.style.height = '0';
                element.style.overflow = 'hidden';
                element.setAttribute('aria-hidden', 'true');
            });
            
            // インフォメーションアイコンと閉じるボタンを探して非表示に
            var icons = document.querySelectorAll(
                'i[class*="info"], i[class*="close"], ' +
                'svg[class*="info"], svg[class*="close"], ' +
                'span[class*="info"], span[class*="close"], ' +
                'button[class*="close"], a[class*="close"]'
            );
            
            icons.forEach(function(icon) {
                icon.style.display = 'none';
                icon.style.visibility = 'hidden';
            });
            
            // 印刷用のスタイルを追加（より強力なバージョン）
            var style = document.createElement('style');
            style.innerHTML = `
                @media print {
                    .alert, .alert-success, .alert-info, .alert-warning, .alert-danger,
                    .notice, .message, .flash-message, .header-message, .notification,
                    .status-message, div[role="alert"], .toast, .banner,
                    .info-message, .success-message, .warning-message, .error-message,
                    .flash, .flash-notice, .flash-success, .flash-error,
                    .message-container, .message-box, .notification-container,
                    .close-button, .close-icon, button.close,
                    div[class*="alert"], div[class*="notice"], div[class*="message"],
                    div[class*="notification"], div[class*="toast"],
                    div[class*="info"], div[class*="success"], div[class*="warning"], div[class*="error"],
                    i[class*="info"], i[class*="close"],
                    svg[class*="info"], svg[class*="close"],
                    span[class*="info"], span[class*="close"],
                    button[class*="close"], a[class*="close"] {
                        display: none !important;
                        visibility: hidden !important;
                        opacity: 0 !important;
                        height: 0 !important;
                        width: 0 !important;
                        overflow: hidden !important;
                        position: absolute !important;
                        top: -9999px !important;
                        left: -9999px !important;
                    }
                    
                    body { 
                        margin: 0 !important; 
                        padding: 0 !important;
                        -webkit-print-color-adjust: exact !important;
                        color-adjust: exact !important;
                    }
                    
                    /* 余分な余白を削除 */
                    * { 
                        margin-top: 0 !important;
                        padding-top: 0 !important;
                    }
                    
                    /* 最初のコンテンツ要素の上部余白を削除 */
                    body > *:first-child {
                        margin-top: 0 !important;
                        padding-top: 0 !important;
                    }
                }
            `;
            document.head.appendChild(style);
            
            // 余分な余白を削除
            document.body.style.margin = '0';
            document.body.style.padding = '0';
            
            // 特定のサイト向けのカスタム処理
            // CrowdWorks特有の要素を探して非表示に
            var cwSpecificElements = document.querySelectorAll(
                '.cw-alert, .cw-notice, .cw-message, .cw-flash, ' +
                '.cw-header-message, .cw-notification, ' +
                '.receipt-header, .receipt-notice, ' +
                'div[class*="cw-alert"], div[class*="cw-notice"], div[class*="cw-message"]'
            );
            
            cwSpecificElements.forEach(function(element) {
                element.style.display = 'none';
                element.style.visibility = 'hidden';
            });
            
            // 最上部の要素を探して、それが通知系の場合は非表示に
            var topElements = Array.from(document.body.children).slice(0, 3);
            topElements.forEach(function(element) {
                var text = element.textContent.toLowerCase();
                if (text.includes('発行しました') || 
                    text.includes('完了') || 
                    text.includes('成功') || 
                    text.includes('通知') ||
                    text.includes('メッセージ')) {
                    element.style.display = 'none';
                    element.style.visibility = 'hidden';
                }
            });
            
            return true;
        """)
        
        # DOMの更新を待つ
        time.sleep(1)
        logger.info("ヘッダー要素を非表示にしました（強化版）")
        return True
    except Exception as e:
        logger.error(f"ヘッダー要素の非表示化に失敗: {str(e)}")
        return False

def extract_receipt_number(driver):
    """ページから領収書番号または請求書番号を抽出する"""
    try:
        # 領収書/請求書番号を含む可能性のある要素を探す
        receipt_number = None
        
        # 方法1: 「領収書番号」または「請求書番号」というラベルを探す
        try:
            label_elements = driver.find_elements(By.XPATH, 
                "//th[contains(text(), '領収書番号') or contains(text(), '請求書番号')] | " + 
                "//td[contains(text(), '領収書番号') or contains(text(), '請求書番号')] | " + 
                "//label[contains(text(), '領収書番号') or contains(text(), '請求書番号')] | " +
                "//div[contains(text(), '領収書番号') or contains(text(), '請求書番号')] | " +
                "//span[contains(text(), '領収書番号') or contains(text(), '請求書番号')]")
            
            for element in label_elements:
                # 親要素または次の兄弟要素を取得
                try:
                    parent = element.find_element(By.XPATH, "./parent::*")
                    next_element = parent.find_element(By.XPATH, "./following-sibling::*[1]")
                    text = next_element.text.strip()
                    
                    # 数字とハイフンを含む文字列を探す
                    if re.search(r'[0-9A-Z-]+', text):
                        receipt_number = re.search(r'[0-9A-Z-]+', text).group()
                        logger.info(f"番号を見つけました（ラベル方式）: {receipt_number}")
                        break
                except:
                    # 次の兄弟要素が見つからない場合は、親要素のテキストから抽出を試みる
                    text = parent.text.strip()
                    match = re.search(r'[：:]\s*([0-9A-Z-]+)', text)
                    if match:
                        receipt_number = match.group(1)
                        logger.info(f"番号を見つけました（親要素テキスト方式）: {receipt_number}")
                        break
        except Exception as e:
            logger.info(f"ラベルから番号の抽出に失敗: {str(e)}")
        
        # 方法2: テーブルから探す
        if not receipt_number:
            try:
                # まずテーブルヘッダーを探す
                headers = driver.find_elements(By.TAG_NAME, "th")
                for header in headers:
                    if '番号' in header.text or 'No' in header.text:
                        # 対応するセルを探す
                        try:
                            header_index = 0
                            all_headers = header.find_elements(By.XPATH, "./parent::*/th")
                            for i, h in enumerate(all_headers):
                                if h == header:
                                    header_index = i
                                    break
                            
                            # 同じ列のセルを取得
                            rows = driver.find_elements(By.TAG_NAME, "tr")
                            for row in rows:
                                cells = row.find_elements(By.TAG_NAME, "td")
                                if len(cells) > header_index:
                                    cell_text = cells[header_index].text.strip()
                                    if re.search(r'[A-Z0-9-]+', cell_text):
                                        receipt_number = re.search(r'[A-Z0-9-]+', cell_text).group()
                                        logger.info(f"番号を見つけました（テーブルヘッダー方式）: {receipt_number}")
                                        break
                        except:
                            pass
                
                # 一般的なテーブルセルから探す
                if not receipt_number:
                    table_cells = driver.find_elements(By.TAG_NAME, "td")
                    for cell in table_cells:
                        text = cell.text.strip()
                        # 数字とハイフンのパターンを探す (例: R-12345678, CW-123456)
                        if re.search(r'[A-Z]-\d+', text) or re.search(r'\d{5,}', text) or re.search(r'CW-\d+', text):
                            receipt_number = text
                            logger.info(f"番号を見つけました（テーブルセル方式）: {receipt_number}")
                            break
            except Exception as e:
                logger.info(f"テーブルから番号の抽出に失敗: {str(e)}")
        
        # 方法3: ページ全体から特定のパターンを探す
        if not receipt_number:
            try:
                page_text = driver.find_element(By.TAG_NAME, "body").text
                # 領収書/請求書番号のパターンを探す
                patterns = [
                    r'領収書番号[：:]\s*([A-Z0-9-]+)',
                    r'請求書番号[：:]\s*([A-Z0-9-]+)',
                    r'領収書[：:]\s*([A-Z0-9-]+)',
                    r'請求書[：:]\s*([A-Z0-9-]+)',
                    r'受領書番号[：:]\s*([A-Z0-9-]+)',
                    r'No[.：:]\s*([A-Z0-9-]+)',
                    r'番号[：:]\s*([A-Z0-9-]+)',
                    r'CW-(\d+)',
                    r'[A-Z]-(\d{5,})'
                ]
                
                for pattern in patterns:
                    match = re.search(pattern, page_text)
                    if match:
                        receipt_number = match.group(1)
                        logger.info(f"番号を見つけました（パターン方式）: {receipt_number}")
                        break
            except Exception as e:
                logger.info(f"ページテキストから番号の抽出に失敗: {str(e)}")
        
        # 方法4: URLから抽出を試みる
        if not receipt_number:
            try:
                current_url = driver.current_url
                # URLから数字の部分を抽出
                url_match = re.search(r'receipt[s]?/(\d+)', current_url) or re.search(r'invoice[s]?/(\d+)', current_url)
                if url_match:
                    receipt_number = url_match.group(1)
                    logger.info(f"番号を見つけました（URL方式）: {receipt_number}")
            except Exception as e:
                logger.info(f"URLから番号の抽出に失敗: {str(e)}")
        
        # 無効な文字を削除
        if receipt_number:
            receipt_number = re.sub(r'[\\/:*?"<>|]', '', receipt_number)
            # 長すぎる場合は短くする
            if len(receipt_number) > 20:
                receipt_number = receipt_number[:20]
        
        return receipt_number
    except Exception as e:
        logger.error(f"番号の抽出中にエラー: {str(e)}")
        return None

def save_as_pdf(driver, index):
    """現在のページをPDFとして保存する（ヘッダー除去強化版）"""
    try:
        # 領収書番号を抽出
        receipt_number = extract_receipt_number(driver)
        
        # ファイル名を生成
        pdf_file_name = generate_pdf_filename(index, receipt_number)
        pdf_path = os.path.join(download_dir, pdf_file_name)
        
        # ヘッダー要素を非表示にする（2回実行して確実に）
        hide_header_elements(driver)
        time.sleep(0.5)
        hide_header_elements(driver)
        
        try:
            # PDFの設定を調整
            pdf = driver.execute_cdp_cmd("Page.printToPDF", {
                "printBackground": True,
                "preferCSSPageSize": True,
                "marginTop": 0,
                "marginBottom": 0,
                "marginLeft": 0,
                "marginRight": 0,
                "scale": 0.9,
                "paperWidth": 8.27,  # A4サイズ
                "paperHeight": 11.69,
                "displayHeaderFooter": False,
                "printBackground": True
            })
            
            # PDFを保存
            with open(pdf_path, "wb") as f:
                f.write(base64.b64decode(pdf["data"]))
            
            logger.info(f"ページをPDFとして保存しました: {pdf_file_name}")
            return pdf_file_name
            
        except Exception as e:
            logger.error(f"CDP方式でのPDF保存に失敗: {str(e)}")
            
            # 代替方法: 印刷ダイアログを使用
            try:
                # 印刷ダイアログを開く前にヘッダー要素を再度非表示に
                hide_header_elements(driver)
                
                driver.execute_script('window.print();')
                time.sleep(3)
                
                print("\n=== PDF保存ダイアログが開いた場合 ===")
                print("PDFとして保存してください")
                print(f"保存先: {download_dir}")
                print(f"ファイル名: {pdf_file_name}")
                
                user_input = input("保存が完了したら「y」を、失敗した場合は「n」を入力してください: ")
                if user_input.lower() == 'y':
                    if os.path.exists(pdf_path):
                        logger.info(f"ユーザーによるPDF保存を確認: {pdf_file_name}")
                        return pdf_file_name
                    else:
                        print(f"ファイル {pdf_file_name} が見つかりません。")
                        print("別の名前で保存した場合は、そのファイル名を入力してください（拡張子含む）:")
                        print("保存していない場合は、Enterキーを押してください。")
                        custom_filename = input().strip()
                        
                        if custom_filename:
                            custom_path = os.path.join(download_dir, custom_filename)
                            if os.path.exists(custom_path):
                                logger.info(f"ユーザーが指定したファイルを確認: {custom_filename}")
                                return custom_filename
                
                return None
                
            except Exception as e2:
                logger.error(f"代替PDF保存方法も失敗: {str(e2)}")
                return None
                
    except Exception as e:
        logger.error(f"PDFとして保存できませんでした: {str(e)}")
        return None

def download_pdf_from_url(driver, pdf_url, index):
    """PDFのURLから直接ダウンロードする"""
    try:
        import requests
        from urllib.parse import urljoin
        
        # 相対URLの場合は絶対URLに変換
        if not pdf_url.startswith('http'):
            pdf_url = urljoin(driver.current_url, pdf_url)
        
        # Cookieを取得
        cookies = driver.get_cookies()
        cookies_dict = {cookie['name']: cookie['value'] for cookie in cookies}
        
        # PDFをダウンロード
        response = requests.get(pdf_url, cookies=cookies_dict)
        
        if response.status_code == 200:
            # ファイル名を生成（タイムスタンプなし）
            pdf_file_name = generate_pdf_filename(index)
            pdf_path = os.path.join(download_dir, pdf_file_name)
            
            # PDFを保存
            with open(pdf_path, "wb") as f:
                f.write(response.content)
            
            logger.info(f"PDFを直接ダウンロードしました: {pdf_file_name}")
            return pdf_file_name
        else:
            logger.error(f"PDFのダウンロードに失敗しました。ステータスコード: {response.status_code}")
            return None
    except Exception as e:
        logger.error(f"PDFの直接ダウンロードに失敗: {str(e)}")
        return None

def parse_arguments():
    """コマンドライン引数を解析する"""
    parser = argparse.ArgumentParser(description='Crowdworksから領収書をダウンロードするスクリプト')
    parser.add_argument('--download-dir', default=None,
                        help='領収書のダウンロード先ディレクトリ')
    parser.add_argument('--log-file', default='receipt_download_manual.log',
                        help='ログファイル名')
    parser.add_argument('--config', default='config.json',
                        help='設定ファイルのパス')
    parser.add_argument('--headless', action='store_true',
                        help='ヘッドレスモードで実行（手動ログイン時は無効）')
    return parser.parse_args()

def load_config(config_path):
    """設定ファイルを読み込む"""
    if os.path.exists(config_path):
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def main():
    """メイン処理"""
    global download_dir
    
    # ログ設定を変更（コンソール出力を無効化）
    setup_logging()
    
    # タイムスタンプ付きのダウンロードディレクトリを作成
    download_dir = create_download_dir()
    
    # 領収書ダウンロード処理の実行
    download_receipts_with_manual_login(download_dir=download_dir)

def wait_for_page_load(driver, timeout=30):
    """ページの読み込みが完了するまで待機する"""
    try:
        # DOMの読み込み完了を待機
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script('return document.readyState') == 'complete'
        )
    except Exception as e:
        logger.warning(f"ページの読み込み待機中にエラー: {str(e)}")
        return False

def safe_click(driver, element):
    """要素を安全にクリックする（複数の方法を試す）"""
    try:
        # 要素が表示されるまでスクロール
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        time.sleep(0.5)  # スクロール後に少し待機
        
        # 方法1: JavaScriptでクリック
        try:
            driver.execute_script("arguments[0].click();", element)
            return True
        except Exception as e:
            logger.debug(f"JavaScriptクリックに失敗: {str(e)}")
        
        # 方法2: 通常のクリック
        try:
            element.click()
            return True
        except Exception as e:
            logger.debug(f"通常クリックに失敗: {str(e)}")
        
        # 方法3: ActionChainsを使用
        try:
            from selenium.webdriver.common.action_chains import ActionChains
            ActionChains(driver).move_to_element(element).click().perform()
            return True
        except Exception as e:
            logger.debug(f"ActionChainsクリックに失敗: {str(e)}")
        
        # 方法4: href属性を取得して直接移動
        try:
            href = element.get_attribute("href")
            if href:
                logger.info(f"直接URLに移動します: {href}")
                driver.get(href)
                return True
        except Exception as e:
            logger.debug(f"直接URL移動に失敗: {str(e)}")
        
        logger.error("すべてのクリック方法が失敗しました")
        return False
    except Exception as e:
        logger.error(f"safe_click処理中にエラー: {str(e)}")
        return False

def display_progress(current, total, description="処理中"):
    """進捗状況を表示する（時間表示なし）"""
    progress = min(current / total * 100, 100)
    bar_length = 40  # バーの長さを調整
    filled_length = int(bar_length * current / total)
    
    bar = '█' * filled_length + '░' * (bar_length - filled_length)
    
    # シンプルな進捗バーのみを表示
    print(f"\r{description}:[{bar}]{progress:.1f}%({current}/{total})", end='          \r')
    
    if current == total:
        print()  # 改行

def move_to_next_page(driver):
    """次のページに移動する（「次へ」リンクのみを使用）"""
    try:
        # 「次へ」リンクを探す（より具体的なXPath）
        try:
            next_link = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((
                    By.XPATH, 
                    "//a[contains(text(), '次へ')] | " +
                    "//a[contains(text(), '次の')] | " +
                    "//a[@rel='next']"
                ))
            )
            logger.info("「次へ」リンクが見つかりました")
            
            # リンクが表示されるまでスクロール
            driver.execute_script("arguments[0].scrollIntoView(true);", next_link)
            time.sleep(1)
            
            # 「次へ」リンクをクリック
            if safe_click(driver, next_link):
                # ページの読み込みを待機
                wait_for_page_load(driver)
                time.sleep(3)
                
                # 領収書リンクの存在を確認
                receipt_links = get_receipt_links(driver)
                if receipt_links:
                    logger.info(f"次のページに {len(receipt_links)} 件の領収書が見つかりました")
                    return True
                else:
                    logger.error("次のページで領収書リンクが見つかりません")
                    return False
            else:
                logger.error("「次へ」リンクのクリックに失敗しました")
                return False
            
        except Exception as e:
            logger.error(f"「次へ」リンクの検索または処理に失敗: {str(e)}")
            return False
        
    except Exception as e:
        logger.error(f"次のページへの移動中にエラー: {str(e)}")
        return False

def process_single_receipt(driver, index, total):
    """単一の領収書を処理する"""
    logger.info(f"領収書 {index}/{total} の処理を開始します")
    display_progress(index-1, total, "領収書ダウンロード")
    
    # 毎回領収書リンクを再取得（stale element referenceを回避）
    try:
        current_links = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, "//a[contains(text(), '領収書')]"))
        )
        
        if index <= len(current_links):
            # index番目の領収書リンクをクリック
            link = current_links[index-1]
            
            # 安全にクリック
            if not safe_click(driver, link):
                logger.error(f"領収書リンク {index} のクリックに失敗しました")
                return False
            
            # ページ読み込みを待機
            wait_for_page_load(driver)
            time.sleep(3)  # 追加の待機時間
            
            # 既に発行済みかどうかを確認（印刷ボタンが存在するか）
            print_button = None
            try:
                print_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, ".print_button, .cw-button_action.print_button"))
                )
            except:
                logger.info("印刷ボタンが見つかりませんでした。未発行の領収書です。")
            
            # 既に発行済みの場合はPDF保存処理へ、そうでなければ発行処理へ
            if print_button:
                # PDFとして保存（印刷ボタンを使わない）
                pdf_file_name = save_as_pdf(driver, index)
                if pdf_file_name:
                    logger.info(f"領収書 {index} を保存しました: {pdf_file_name}")
                    display_progress(index, total, "領収書ダウンロード")
                    return True
                else:
                    # PDF保存に失敗した場合はスクリーンショットを取る
                    logger.info("PDFとして保存に失敗しました。スクリーンショットを取ります...")
                    screenshot_path = os.path.join(download_dir, f"領収書_{index:02d}.png")
                    driver.save_screenshot(screenshot_path)
                    logger.info(f"領収書 {index} をスクリーンショットとして保存しました: {screenshot_path}")
                    display_progress(index, total, "領収書ダウンロード")
                    return True
            else:
                # 未発行の領収書の場合、発行ボタンを探して処理
                try:
                    # まず「プレビューで内容を確認する」ボタンを探す
                    preview_button = None
                    try:
                        preview_button = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, "//input[@value='プレビューで内容を確認する'] | //button[contains(text(), 'プレビューで内容を確認する')]"))
                        )
                        logger.info("「プレビューで内容を確認する」ボタンを見つけました。クリックします。")
                        
                        if safe_click(driver, preview_button):
                            wait_for_page_load(driver)
                            time.sleep(2)
                            # プレビュー画面でヘッダーを非表示に
                            hide_header_elements(driver)
                            logger.info("プレビュー画面に移動しました。")
                    except Exception as e:
                        logger.info(f"「プレビューで内容を確認する」ボタンが見つからないか、クリックに失敗しました: {str(e)}")
                        logger.info("プレビュー画面をスキップして発行処理を続行します。")
                    
                    # 次に「この内容で発行する」ボタンを探す
                    issue_button = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//input[@value='この内容で発行する'] | //button[contains(text(), 'この内容で発行する')]"))
                    )
                    logger.info("「この内容で発行する」ボタンを見つけました。クリックします。")
                    
                    if safe_click(driver, issue_button):
                        wait_for_page_load(driver)
                        time.sleep(2)
                        # 発行直後にヘッダーを非表示に
                        hide_header_elements(driver)
                        
                        # 確認ダイアログの「はい」ボタンを探して処理
                        try:
                            yes_button = WebDriverWait(driver, 5).until(
                                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'はい')] | //input[@value='はい'] | //a[contains(text(), 'はい')] | //div[contains(@class, 'dialog')]//button[1]"))
                            )
                            logger.info("確認ダイアログの「はい」ボタンを見つけました。クリックします。")
                            
                            if safe_click(driver, yes_button):
                                wait_for_page_load(driver)
                                time.sleep(3)
                                # 発行完了後にもヘッダーを非表示に
                                hide_header_elements(driver)
                                
                                # 発行後のページをPDFとして保存
                                pdf_file_name = save_as_pdf(driver, index)
                                if pdf_file_name:
                                    logger.info(f"領収書 {index} を保存しました: {pdf_file_name}")
                                    return True
                                else:
                                    # PDF保存に失敗した場合はスクリーンショットを取る
                                    logger.info("PDFとして保存に失敗しました。スクリーンショットを取ります...")
                                    screenshot_path = os.path.join(download_dir, f"領収書_{index:02d}.png")
                                    driver.save_screenshot(screenshot_path)
                                    logger.info(f"領収書 {index} をスクリーンショットとして保存しました: {screenshot_path}")
                                    return True
                        except Exception as e:
                            logger.error(f"確認ダイアログの「はい」ボタン処理中にエラー: {str(e)}")
                            # JavaScriptでの処理を試みる
                            try:
                                driver.execute_script("""
                                    // 確認ダイアログの「はい」ボタンを探して自動的にクリック
                                    var yesButtons = document.querySelectorAll('button, input[type="button"], input[type="submit"], a.button');
                                    for (var i = 0; i < yesButtons.length; i++) {
                                        var btn = yesButtons[i];
                                        if (btn.textContent.includes('はい') || btn.value === 'はい') {
                                            btn.click();
                                            return true;
                                        }
                                    }
                                    // 最初のボタンをクリック（多くの場合「はい」が最初）
                                    var firstButton = document.querySelector('.dialog button, .modal button, .confirm button');
                                    if (firstButton) {
                                        firstButton.click();
                                        return true;
                                    }
                                    return false;
                                """)
                                time.sleep(3)
                                # JavaScript実行後にもヘッダーを非表示に
                                hide_header_elements(driver)
                            except Exception as js_error:
                                logger.error(f"JavaScriptによる確認ダイアログ処理に失敗: {str(js_error)}")

                        # 発行完了後の最終確認としてヘッダーを非表示に
                        hide_header_elements(driver)
                        
                        # 発行後のページをPDFとして保存
                        pdf_file_name = save_as_pdf(driver, index)
                        if pdf_file_name:
                            logger.info(f"領収書 {index} を保存しました: {pdf_file_name}")
                            return True
                        else:
                            # PDF保存に失敗した場合はスクリーンショットを取る
                            logger.info("PDFとして保存に失敗しました。スクリーンショットを取ります...")
                            screenshot_path = os.path.join(download_dir, f"領収書_{index:02d}.png")
                            driver.save_screenshot(screenshot_path)
                            logger.info(f"領収書 {index} をスクリーンショットとして保存しました: {screenshot_path}")
                            return True
                except Exception as e:
                    logger.error(f"発行ボタンの処理中にエラー: {str(e)}")
            
            # 手動操作を求める
            print("=== 自動処理に失敗しました ===")
            print("手動で領収書を発行・保存してください。")
            input("操作が完了したら、Enterキーを押して続行してください...")
            return True
        else:
            logger.error(f"領収書リンク {index} が見つかりません（リンク数: {len(current_links)}）")
            return False
    except Exception as e:
        logger.error(f"領収書 {index} の処理中にエラー: {str(e)}")
        return False

def handle_receipt_error(driver, error, index, page_num):
    """領収書処理中のエラーを処理する"""
    error_msg = str(error) if str(error) else "不明なエラー（エラーメッセージなし）"
    logger.error(f"領収書 {index} の処理中にエラーが発生: {error_msg}")
    
    # スクリーンショットを保存（タイムスタンプなし）
    screenshot_path = os.path.join(download_dir, f"error_screenshot_page{page_num}_receipt{index}.png")
    driver.save_screenshot(screenshot_path)
    logger.info(f"エラー時のスクリーンショットを保存: {screenshot_path}")
    
    # 手動介入を求める
    print(f"=== 領収書 {index} の処理中にエラーが発生しました ===")
    print("エラー内容:", error_msg)
    print("1: 再試行する")
    print("2: この領収書をスキップして次に進む")
    print("3: 処理を中止する")
    choice = input("選択してください (1/2/3): ")
    
    if choice == "1":
        # 現在の領収書を再処理
        logger.info("ユーザーが再試行を選択しました")
        return True
    elif choice == "3":
        logger.info("ユーザーにより処理が中止されました")
        return False
    else:
        # 一覧ページに戻る
        logger.info("ユーザーがスキップを選択しました")
        return go_back_to_list_page(driver, page_num)

def find_pdf_url(driver):
    """ページ内のPDF URLを探す"""
    pdf_url = None
    
    try:
        # 埋め込みPDFを検索
        pdf_embed = driver.find_element(By.CSS_SELECTOR, "embed[type='application/pdf'], object[type='application/pdf'], iframe[src$='.pdf']")
        pdf_url = pdf_embed.get_attribute("src")
        logger.info(f"埋め込みPDFを見つけました: {pdf_url}")
        return pdf_url
    except:
        logger.info("埋め込みPDFは見つかりませんでした")
    
    try:
        # PDFへのリンクを検索
        pdf_link = driver.find_element(By.CSS_SELECTOR, "a[href$='.pdf'], a[href*='pdf'], a[download]")
        pdf_url = pdf_link.get_attribute("href")
        logger.info(f"PDFへのリンクを見つけました: {pdf_url}")
        return pdf_url
    except:
        logger.info("PDFへのリンクは見つかりませんでした")
    
    return None

def setup_logging():
    """ログ出力を設定する（ファイルのみに出力し、コンソールには出力しない）"""
    global logger
    
    # ロガーの設定
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)
    
    # ファイルハンドラの設定
    file_handler = logging.FileHandler("receipt_download_manual.log")
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    
    # ロガーにハンドラを追加
    logger.handlers = []  # 既存のハンドラをクリア
    logger.addHandler(file_handler)
    
    # コンソール出力を無効化
    logger.propagate = False

if __name__ == "__main__":
    main()