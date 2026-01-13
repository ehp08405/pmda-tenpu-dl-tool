#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PMDA 添付文書 統合ツール (GUI版 - 中止機能付き)
==================================================
PMDAの添付文書一覧ページから情報を収集し(Step 1)、
PDFファイルをダウンロードする(Step 2)ための統合ツールです。

主な機能:
1. [一覧取得タブ]: 日付指定でスクレイピングしCSVを作成
2. [ダウンロードタブ]: CSVを読み込みPDFを一括ダウンロード
   (病院内機器リストによるフィルタリング機能付き)

変更点:
- 一覧取得タブに中止ボタンを追加
- 収集中・ダウンロード中に処理を中断可能

使い方:
    python pmda_tool_with_settings.py
"""

import requests
from bs4 import BeautifulSoup
import csv
from datetime import datetime, timedelta
import time
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import json

# ==========================================
# 定数・設定
# ==========================================
# PyInstallerでEXE化した場合とスクリプト実行時の両方に対応
if getattr(sys, 'frozen', False):
    # EXE化されている場合: 実行ファイルと同じディレクトリ
    CONFIG_FILE = os.path.join(os.path.dirname(sys.executable), "pmda_config.json")
else:
    # スクリプト実行時: スクリプトと同じディレクトリ
    CONFIG_FILE = os.path.join(os.path.dirname(__file__), "pmda_config.json")

# デフォルト設定値
DEFAULT_SETTINGS = {
    "base_url": "https://www.info.pmda.go.jp/ysearch/tenpulist.jsp?DATE=",
    "detail_base_url": "https://www.info.pmda.go.jp",
    "default_output_dir": "./output",
    "default_wait_time": 0.5,
    "scrape_wait_time": 0.3,
    "approval_number_wait_time": 0.5,
    "window_width": 700,
    "window_height": 680,
    "window_x": None,
    "window_y": None
}

# HTTP設定
HTTP_TIMEOUT = 30
HTTP_DOWNLOAD_TIMEOUT = 60
HTTP_CHUNK_SIZE = 8192
DEFAULT_USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'

# URLパス構成
YGO_PACK_PREFIX = "ygo/pack"
YGO_PDF_PREFIX = "ygo/pdf"

# CSV設定
CSV_ENCODING = 'utf-8-sig'

# ファイル名設定
PDF_EXTENSION = '.pdf'
CSV_FILENAME_PREFIX = 'pmda_list_'
CSV_FILENAME_FORMAT = '%Y%m%d_%H%M%S'

# セクション名
SECTION_LISTED = '掲載'
SECTION_DELETED = '削除'

# 待機時間制限
MIN_WAIT_TIME = 0
MAX_WAIT_TIME = 10

# グローバル設定（アプリ起動時に読み込み）
class Settings:
    def __init__(self):
        """デフォルト設定で初期化"""
        self._apply_settings(DEFAULT_SETTINGS)

    def _apply_settings(self, settings_dict):
        """設定辞書から属性を設定"""
        for key, value in settings_dict.items():
            setattr(self, key, value)

    def load(self):
        """設定ファイルから読み込み"""
        if not os.path.exists(CONFIG_FILE):
            return

        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # デフォルト値をベースに、ファイルの値でオーバーライド
                merged = {**DEFAULT_SETTINGS, **data}
                self._apply_settings(merged)
        except Exception as e:
            print(f"設定読み込みエラー: {e}")

    def save(self):
        """設定ファイルに保存"""
        try:
            data = {key: getattr(self, key) for key in DEFAULT_SETTINGS.keys()}
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"設定保存エラー: {e}")

    def reset(self):
        """デフォルト値に戻す"""
        self._apply_settings(DEFAULT_SETTINGS)

# グローバル設定インスタンス
app_settings = Settings()

# 後方互換性のため
BASE_URL = app_settings.base_url
DETAIL_BASE_URL = app_settings.detail_base_url
DEFAULT_OUTPUT_DIR = app_settings.default_output_dir
DEFAULT_WAIT_TIME = app_settings.default_wait_time

# ==========================================
# 共通ロジック (Helper Functions)
# ==========================================


def convert_detail_url_to_pdf(detail_url):
    """詳細URLからPDF URLを生成"""
    if not detail_url:
        return ""
    path = detail_url.replace(app_settings.detail_base_url, "")
    parts = path.strip("/").split("/")
    ygo, pack = YGO_PACK_PREFIX.split("/")
    if len(parts) >= 4 and parts[0] == ygo and parts[1] == pack:
        company_id = parts[2]
        doc_id = parts[3]
        return f"{app_settings.detail_base_url}/{YGO_PDF_PREFIX}/{company_id}_{doc_id}/"
    return ""

def convert_detail_url_to_body_url(detail_url):
    """詳細URLからbody URLを生成"""
    if not detail_url:
        return ""
    path = detail_url.replace(app_settings.detail_base_url, "")
    parts = path.strip("/").split("/")
    ygo, pack = YGO_PACK_PREFIX.split("/")
    if len(parts) >= 4 and parts[0] == ygo and parts[1] == pack:
        doc_id = parts[3]
        return f"{detail_url.rstrip('/')}/{doc_id}?view=body"
    return ""

def fetch_approval_number(detail_url, log_callback=None):
    """詳細ページのbody URLから認証番号または承認番号を取得"""
    body_url = convert_detail_url_to_body_url(detail_url)
    if not body_url:
        return "", ""

    try:
        headers = {'User-Agent': DEFAULT_USER_AGENT}
        response = requests.get(body_url, headers=headers, timeout=HTTP_TIMEOUT)
        response.encoding = 'utf-8'
        soup = BeautifulSoup(response.text, 'html.parser')

        approval_no = ""
        certification_no = ""

        # 認証番号を探す
        cert_header = soup.find('h3', class_='section_header', string='認証番号')
        if cert_header:
            next_div = cert_header.find_next('div')
            if next_div:
                certification_no = next_div.get_text(strip=True)

        # 承認番号を探す
        approval_header = soup.find('h3', class_='section_header', string='承認番号')
        if approval_header:
            next_div = approval_header.find_next('div')
            if next_div:
                approval_no = next_div.get_text(strip=True)

        return approval_no, certification_no

    except Exception as e:
        if log_callback:
            log_callback(f"    番号取得エラー: {e}")
        return "", ""

def scrape_date(date_str, fetch_numbers=False, log_callback=None, cancel_check=None):
    """指定日付のページをスクレイピング"""
    url = app_settings.base_url + date_str.replace("-", "")
    results = []

    try:
        headers = {'User-Agent': DEFAULT_USER_AGENT}
        response = requests.get(url, headers=headers, timeout=HTTP_TIMEOUT)
        response.encoding = 'utf-8'
        soup = BeautifulSoup(response.text, 'html.parser')

        tables = soup.find_all('table')
        current_section = None

        for table in tables:
            prev_h2 = table.find_previous('h2')
            if prev_h2:
                if '掲載分' in prev_h2.get_text():
                    current_section = SECTION_LISTED
                elif '削除分' in prev_h2.get_text():
                    current_section = SECTION_DELETED

            if current_section is None:
                continue

            rows = table.find_all('tr')
            for row in rows:
                # 中止チェック
                if cancel_check and not cancel_check():
                    return results

                cells = row.find_all('td')
                if len(cells) >= 3:
                    name_cell = cells[0]
                    name_link = name_cell.find('a')
                    if name_link:
                        product_name = name_link.get_text(strip=True)
                        detail_url = app_settings.detail_base_url + name_link.get('href', '')
                    else:
                        product_name = name_cell.get_text(strip=True)
                        detail_url = ''

                    company = cells[1].get_text(strip=True)
                    company = company.replace('製造販売／', '')
                    reason = cells[2].get_text(strip=True)
                    pdf_url = convert_detail_url_to_pdf(detail_url)

                    approval_no = ""
                    certification_no = ""
                    if fetch_numbers and detail_url and current_section == SECTION_LISTED:
                        # 中止チェック
                        if cancel_check and not cancel_check():
                            return results
                        if log_callback:
                            log_callback(f"    番号取得中: {product_name[:30]}...")
                        approval_no, certification_no = fetch_approval_number(detail_url, log_callback)
                        time.sleep(DEFAULT_SETTINGS["approval_number_wait_time"])

                    if product_name and product_name != '販売名':
                        results.append({
                            '日付': date_str,
                            '区分': current_section,
                            '販売名': product_name,
                            '企業名': company,
                            '理由': reason,
                            '承認番号': approval_no,
                            '認証番号': certification_no,
                            '詳細URL': detail_url,
                            'PDF_URL': pdf_url
                        })
        return results
    except Exception as e:
        if log_callback:
            log_callback(f"  エラー: {e}")
        return []

def extract_doc_id_from_url(url):
    """URLからドキュメントIDを抽出"""
    parts = url.strip("/").split("/")
    if parts:
        return parts[-1]
    return None

def download_file(url, save_path):
    """ファイルをダウンロード"""
    try:
        headers = {'User-Agent': DEFAULT_USER_AGENT}
        response = requests.get(url, headers=headers, timeout=HTTP_DOWNLOAD_TIMEOUT, stream=True)
        if response.status_code == 200:
            with open(save_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=HTTP_CHUNK_SIZE):
                    f.write(chunk)
            return True, None
        else:
            return False, f"HTTPエラー: {response.status_code}"
    except Exception as e:
        return False, str(e)

def load_hospital_device_list(csv_path):
    """病院内機器リストCSVを読み込み、承認番号・認証番号のセットを返す"""
    approval_numbers = set()
    certification_numbers = set()

    with open(csv_path, 'r', encoding=CSV_ENCODING) as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames
        approval_col = None
        certification_col = None

        for col in fieldnames:
            if '承認番号' in col:
                approval_col = col
            if '認証番号' in col:
                certification_col = col

        for row in reader:
            if approval_col and row.get(approval_col):
                approval_numbers.add(row[approval_col].strip())
            if certification_col and row.get(certification_col):
                certification_numbers.add(row[certification_col].strip())

    return approval_numbers, certification_numbers


# ==========================================
# 共通基底クラス
# ==========================================
class BaseTaskTab(ttk.Frame):
    """タスク実行タブの基底クラス（ログ、プログレスバー、ボタン制御の共通機能）"""

    def log(self, msg):
        """ログメッセージを追加（スレッドセーフ）"""
        def _log():
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.insert(tk.END, msg + "\n")
            self.log_text.see(tk.END)
            self.log_text.configure(state=tk.DISABLED)

        # GUIスレッドで実行
        try:
            self.winfo_toplevel().after(0, _log)
        except:
            # フォールバック（初期化中などの場合）
            _log()

    def clear_log(self):
        """ログをクリア"""
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def set_running_state(self, is_running):
        """実行中の状態を設定（ボタンの有効/無効を切り替え）"""
        if is_running:
            self.run_button.configure(state=tk.DISABLED)
            self.cancel_button.configure(state=tk.NORMAL)
        else:
            self.run_button.configure(state=tk.NORMAL)
            self.cancel_button.configure(state=tk.DISABLED)

    def cancel_task(self):
        """タスクの中止（サブクラスでオーバーライド可能）"""
        self.is_running = False
        self.log("中止を要求しました...")


# ==========================================
# Step1: 一覧取得タブ (ScrapeTab)
# ==========================================
class ScrapeTab(BaseTaskTab):
    def __init__(self, parent, app):
        super().__init__(parent, padding="10")
        self.app = app
        self.is_running = False  # 中止制御用フラグ

        # 日付モード選択
        mode_frame = ttk.LabelFrame(self, text="取得モード", padding="10")
        mode_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.date_mode = tk.StringVar(value="single")
        ttk.Radiobutton(mode_frame, text="単一日付", variable=self.date_mode, 
                       value="single", command=self.toggle_date_mode).pack(anchor=tk.W)
        ttk.Radiobutton(mode_frame, text="期間指定", variable=self.date_mode, 
                       value="period", command=self.toggle_date_mode).pack(anchor=tk.W)

        # 日付範囲
        period_frame = ttk.LabelFrame(self, text="取得日付", padding="10")
        period_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 単一日付用
        self.single_frame = ttk.Frame(period_frame)
        ttk.Label(self.single_frame, text="日付:").pack(side=tk.LEFT)
        self.single_year = tk.StringVar(value=str(datetime.now().year))
        self.single_month = tk.StringVar(value=str(datetime.now().month))
        self.single_day = tk.StringVar(value=str(datetime.now().day))
        ttk.Entry(self.single_frame, textvariable=self.single_year, width=6).pack(side=tk.LEFT, padx=(5, 0))
        ttk.Label(self.single_frame, text="年").pack(side=tk.LEFT)
        ttk.Entry(self.single_frame, textvariable=self.single_month, width=4).pack(side=tk.LEFT, padx=(5, 0))
        ttk.Label(self.single_frame, text="月").pack(side=tk.LEFT)
        ttk.Entry(self.single_frame, textvariable=self.single_day, width=4).pack(side=tk.LEFT, padx=(5, 0))
        ttk.Label(self.single_frame, text="日").pack(side=tk.LEFT)
        
        # 期間指定用
        self.period_frame = ttk.Frame(period_frame)
        start_f = ttk.Frame(self.period_frame)
        start_f.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(start_f, text="開始日:").pack(side=tk.LEFT)
        self.start_year = tk.StringVar(value=str(datetime.now().year))
        self.start_month = tk.StringVar(value=str(datetime.now().month))
        self.start_day = tk.StringVar(value="1")
        ttk.Entry(start_f, textvariable=self.start_year, width=6).pack(side=tk.LEFT, padx=(5, 0))
        ttk.Label(start_f, text="年").pack(side=tk.LEFT)
        ttk.Entry(start_f, textvariable=self.start_month, width=4).pack(side=tk.LEFT, padx=(5, 0))
        ttk.Label(start_f, text="月").pack(side=tk.LEFT)
        ttk.Entry(start_f, textvariable=self.start_day, width=4).pack(side=tk.LEFT, padx=(5, 0))
        ttk.Label(start_f, text="日").pack(side=tk.LEFT)
        
        end_f = ttk.Frame(self.period_frame)
        end_f.pack(fill=tk.X)
        ttk.Label(end_f, text="終了日:").pack(side=tk.LEFT)
        self.end_year = tk.StringVar(value=str(datetime.now().year))
        self.end_month = tk.StringVar(value=str(datetime.now().month))
        self.end_day = tk.StringVar(value=str(datetime.now().day))
        ttk.Entry(end_f, textvariable=self.end_year, width=6).pack(side=tk.LEFT, padx=(5, 0))
        ttk.Label(end_f, text="年").pack(side=tk.LEFT)
        ttk.Entry(end_f, textvariable=self.end_month, width=4).pack(side=tk.LEFT, padx=(5, 0))
        ttk.Label(end_f, text="月").pack(side=tk.LEFT)
        ttk.Entry(end_f, textvariable=self.end_day, width=4).pack(side=tk.LEFT, padx=(5, 0))
        ttk.Label(end_f, text="日").pack(side=tk.LEFT)
        
        # 初期表示は単一日付
        self.single_frame.pack(fill=tk.X)

        # オプション
        opt_frame = ttk.LabelFrame(self, text="オプション", padding="10")
        opt_frame.pack(fill=tk.X, pady=(0, 10))
        self.fetch_num_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_frame, text="承認番号・認証番号も取得する (取得時間が長くなります)", 
                       variable=self.fetch_num_var).pack(anchor=tk.W)
        ttk.Label(opt_frame, text="※番号取得はPMDAサーバー負荷を考慮し、0.5秒間隔で実行されます", 
                 foreground="gray", font=("", 8)).pack(anchor=tk.W)

        # 出力設定
        out_frame = ttk.LabelFrame(self, text="出力設定", padding="10")
        out_frame.pack(fill=tk.X, pady=(0, 10))
        
        out_f = ttk.Frame(out_frame)
        out_f.pack(fill=tk.X)
        ttk.Label(out_f, text="保存先:").pack(side=tk.LEFT)
        self.out_dir_var = tk.StringVar(value=app_settings.default_output_dir)
        ttk.Entry(out_f, textvariable=self.out_dir_var, width=45).pack(side=tk.LEFT, padx=(5, 5))
        ttk.Button(out_f, text="参照...", command=self.browse_dir).pack(side=tk.LEFT)

        # 実行ボタン
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        self.run_button = ttk.Button(btn_frame, text="収集開始", command=self.run_scrape)
        self.run_button.pack(side=tk.LEFT, padx=(0, 10))
        self.cancel_button = ttk.Button(btn_frame, text="中止", command=self.cancel_scrape, state=tk.DISABLED)
        self.cancel_button.pack(side=tk.LEFT)

        # ログ
        log_group = ttk.LabelFrame(self, text="ログ", padding="5")
        log_group.pack(fill=tk.BOTH, expand=True)
        self.log_text = tk.Text(log_group, height=8, state=tk.DISABLED)
        sb = ttk.Scrollbar(log_group, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=sb.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

    def toggle_date_mode(self):
        """日付モードの切り替え"""
        if self.date_mode.get() == "single":
            self.period_frame.pack_forget()
            self.single_frame.pack(fill=tk.X)
        else:
            self.single_frame.pack_forget()
            self.period_frame.pack(fill=tk.X)

    def browse_dir(self):
        d = filedialog.askdirectory(initialdir=self.out_dir_var.get())
        if d: self.out_dir_var.set(d)

    def run_scrape(self):
        try:
            mode = self.date_mode.get()
            
            if mode == "single":
                # 単一日付モード
                s_y = int(self.single_year.get())
                s_m = int(self.single_month.get())
                s_d = int(self.single_day.get())
                start = datetime(s_y, s_m, s_d)
                end = start
            else:
                # 期間指定モード
                s_y = int(self.start_year.get())
                s_m = int(self.start_month.get())
                s_d = int(self.start_day.get())
                e_y = int(self.end_year.get())
                e_m = int(self.end_month.get())
                e_d = int(self.end_day.get())
                start = datetime(s_y, s_m, s_d)
                end = datetime(e_y, e_m, e_d)
        except Exception as e:
            messagebox.showerror("エラー", f"日付が不正です: {e}")
            return
        
        if start > end:
            messagebox.showerror("エラー", "開始日が終了日より後になっています")
            return
        
        out_dir = self.out_dir_var.get().strip()
        fetch_nums = self.fetch_num_var.get()

        self.is_running = True
        self.set_running_state(True)
        self.clear_log()

        thread = threading.Thread(target=self.scrape_thread,
                                  args=(start, end, out_dir, fetch_nums))
        thread.daemon = True
        thread.start()

    def cancel_scrape(self):
        """中止処理"""
        self.cancel_task()

    def _generate_date_range(self, start_date, end_date):
        """日付範囲のリストを生成"""
        dates = []
        cur = start_date
        while cur <= end_date:
            dates.append(cur.strftime("%Y-%m-%d"))
            cur += timedelta(days=1)
        return dates

    def _log_scrape_start(self, start_date, end_date, dates, fetch_nums):
        """収集開始のログを出力"""
        self.log("=== 収集開始 ===")
        if start_date == end_date:
            self.log(f"日付: {start_date.strftime('%Y-%m-%d')}")
        else:
            self.log(f"期間: {start_date.strftime('%Y-%m-%d')} ~ {end_date.strftime('%Y-%m-%d')}")
            self.log(f"対象日数: {len(dates)}日")
        if fetch_nums:
            self.log("※承認番号・認証番号も取得します（時間がかかります）")
        self.log("")

    def _collect_data(self, dates, fetch_nums):
        """データを収集"""
        all_results = []
        total = len(dates)

        for i, date_str in enumerate(dates, 1):
            if not self.is_running:
                break

            self.log(f"[{i}/{total}] {date_str} 収集中...")
            results = scrape_date(date_str, fetch_nums, self.log, lambda: self.is_running)
            all_results.extend(results)
            self.log(f"  → {len(results)}件取得")

            time.sleep(app_settings.scrape_wait_time)

        return all_results

    def _handle_scrape_completion(self, all_results, out_dir):
        """収集完了時の処理"""
        if not self.is_running:
            self.log("")
            self.log("=== 中断されました ===")
            if all_results:
                self.log(f"中断までに {len(all_results)}件 取得しました")
                if messagebox.askyesno("確認", f"{len(all_results)}件のデータを保存しますか?"):
                    self.save_results(all_results, out_dir)
        else:
            self.log("")
            self.log(f"=== 収集完了: 合計 {len(all_results)}件 ===")

            if all_results:
                output_path = self.save_results(all_results, out_dir)
                messagebox.showinfo("完了", f"収集が完了しました。\n\n取得件数: {len(all_results)}件\n保存先: {output_path}")
                self.app.set_download_csv_path(output_path)
            else:
                messagebox.showinfo("完了", "データが見つかりませんでした")

    def scrape_thread(self, start_date, end_date, out_dir, fetch_nums):
        """スクレイピング処理のメインスレッド"""
        try:
            os.makedirs(out_dir, exist_ok=True)

            dates = self._generate_date_range(start_date, end_date)
            self._log_scrape_start(start_date, end_date, dates, fetch_nums)

            all_results = self._collect_data(dates, fetch_nums)
            self._handle_scrape_completion(all_results, out_dir)

        except Exception as e:
            self.log(f"エラー: {e}")
            messagebox.showerror("エラー", str(e))
        finally:
            self.is_running = False
            self.set_running_state(False)

    def save_results(self, results, out_dir):
        """結果をCSVに保存"""
        ts = datetime.now().strftime(CSV_FILENAME_FORMAT)
        output_path = os.path.join(out_dir, f"{CSV_FILENAME_PREFIX}{ts}.csv")

        fieldnames = ['日付', '区分', '販売名', '企業名', '理由', '承認番号', '認証番号', '詳細URL', 'PDF_URL']
        with open(output_path, 'w', encoding=CSV_ENCODING, newline='') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(results)

        self.log(f"保存完了: {output_path}")
        return output_path


# ==========================================
# Step2: PDFダウンロードタブ (DownloadTab)
# ==========================================
class DownloadTab(BaseTaskTab):
    def __init__(self, parent):
        super().__init__(parent, padding="10")
        self.is_running = False  # 中止制御用フラグ

        # CSV選択
        csv_frame = ttk.LabelFrame(self, text="Step 1で作成したCSVファイル", padding="10")
        csv_frame.pack(fill=tk.X, pady=(0, 10))

        c_frame = ttk.Frame(csv_frame)
        c_frame.pack(fill=tk.X)
        ttk.Label(c_frame, text="CSVパス:").pack(side=tk.LEFT)
        self.csv_path_var = tk.StringVar()
        self.csv_entry = ttk.Entry(c_frame, textvariable=self.csv_path_var, width=50)
        self.csv_entry.pack(side=tk.LEFT, padx=(5, 5))
        ttk.Button(c_frame, text="参照...", command=self.browse_csv).pack(side=tk.LEFT)

        # フィルタリング
        filter_frame = ttk.LabelFrame(self, text="病院内機器リストによるフィルタリング(オプション)", padding="10")
        filter_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.use_filter_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(filter_frame, text="病院内機器リストでフィルタリングする",
                       variable=self.use_filter_var, command=self.toggle_filter).pack(anchor=tk.W, pady=(0, 5))

        h_frame = ttk.Frame(filter_frame)
        h_frame.pack(fill=tk.X)
        ttk.Label(h_frame, text="機器リストCSV:").pack(side=tk.LEFT)
        self.hosp_csv_var = tk.StringVar()
        self.hosp_entry = ttk.Entry(h_frame, textvariable=self.hosp_csv_var, width=40, state=tk.NORMAL)
        self.hosp_entry.pack(side=tk.LEFT, padx=(5, 5))
        self.hosp_btn = ttk.Button(h_frame, text="参照...", command=self.browse_hosp_csv, state=tk.NORMAL)
        self.hosp_btn.pack(side=tk.LEFT)

        self.hosp_note_label = ttk.Label(filter_frame, text="※「承認番号」「認証番号」列が必要です", foreground="gray")
        self.hosp_note_label.pack(anchor=tk.W)

        # 出力とオプション
        out_frame = ttk.LabelFrame(self, text="出力設定・オプション", padding="10")
        out_frame.pack(fill=tk.X, pady=(0, 10))
        
        d_frame = ttk.Frame(out_frame)
        d_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(d_frame, text="保存先:").pack(side=tk.LEFT)
        self.out_dir_var = tk.StringVar(value=os.path.join(app_settings.default_output_dir, "downloads"))
        ttk.Entry(d_frame, textvariable=self.out_dir_var, width=45).pack(side=tk.LEFT, padx=(5, 5))
        ttk.Button(d_frame, text="参照...", command=self.browse_out_dir).pack(side=tk.LEFT)
        
        o_frame = ttk.Frame(out_frame)
        o_frame.pack(fill=tk.X)
        self.skip_exist_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(o_frame, text="既存ファイルをスキップ", variable=self.skip_exist_var).pack(side=tk.LEFT)

        # 実行ボタン
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        self.run_button = ttk.Button(btn_frame, text="ダウンロード実行", command=self.run_download)
        self.run_button.pack(side=tk.LEFT, padx=(0, 10))
        self.cancel_button = ttk.Button(btn_frame, text="中止", command=self.cancel_download, state=tk.DISABLED)
        self.cancel_button.pack(side=tk.LEFT)

        # ログ
        log_group = ttk.LabelFrame(self, text="ログ", padding="5")
        log_group.pack(fill=tk.BOTH, expand=True)
        self.log_text = tk.Text(log_group, height=8, state=tk.DISABLED)
        sb = ttk.Scrollbar(log_group, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=sb.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

    def toggle_filter(self):
        """フィルタリングのON/OFFを切り替え"""
        is_enabled = self.use_filter_var.get()
        st = tk.NORMAL if is_enabled else tk.DISABLED

        self.hosp_entry.configure(state=st)
        self.hosp_btn.configure(state=st)

    def browse_csv(self):
        f = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
        if f: self.csv_path_var.set(f)

    def browse_hosp_csv(self):
        f = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
        if f: self.hosp_csv_var.set(f)

    def browse_out_dir(self):
        d = filedialog.askdirectory(initialdir=self.out_dir_var.get())
        if d: self.out_dir_var.set(d)

    def run_download(self):
        pmda_csv = self.csv_path_var.get().strip()
        out_dir = self.out_dir_var.get().strip()
        
        if not pmda_csv or not os.path.exists(pmda_csv):
            messagebox.showerror("エラー", "PMDA CSVファイルが存在しません")
            return
        
        use_filter = self.use_filter_var.get()
        hosp_csv = self.hosp_csv_var.get().strip()
        if use_filter and (not hosp_csv or not os.path.exists(hosp_csv)):
            messagebox.showerror("エラー", "病院機器リストCSVを選択してください")
            return

        wait = app_settings.default_wait_time

        self.is_running = True
        self.set_running_state(True)
        self.clear_log()

        thread = threading.Thread(target=self.download_thread,
                                  args=(pmda_csv, out_dir, use_filter, hosp_csv, self.skip_exist_var.get(), wait))
        thread.daemon = True
        thread.start()

    def cancel_download(self):
        """中止処理"""
        self.cancel_task()

    def _load_filter_lists(self, use_filter, hosp_csv):
        """病院リストからフィルタリング用のセットを読み込み"""
        target_approvals, target_certs = set(), set()
        if use_filter:
            self.log("病院リスト読み込み中...")
            try:
                target_approvals, target_certs = load_hospital_device_list(hosp_csv)
                self.log(f"  対象: 承認番号{len(target_approvals)}件, 認証番号{len(target_certs)}件")
            except Exception as e:
                self.log(f"リスト読み込みエラー: {e}")
                raise
        return target_approvals, target_certs

    def _load_and_filter_targets(self, csv_file, use_filter, target_approvals, target_certs):
        """CSVを読み込み、フィルタリング後の対象を返す"""
        with open(csv_file, 'r', encoding=CSV_ENCODING) as f:
            reader = csv.DictReader(f)
            rows = list(reader)

        targets = [r for r in rows if r.get('区分') == SECTION_LISTED and r.get('PDF_URL')]

        if use_filter:
            filtered = []
            for r in targets:
                an = r.get('承認番号', '').strip()
                cn = r.get('認証番号', '').strip()
                if (an and an in target_approvals) or (cn and cn in target_certs):
                    filtered.append(r)
            self.log(f"全{len(targets)}件中、病院リストと一致: {len(filtered)}件")
            targets = filtered

        return targets

    def _download_files(self, targets, pdf_dir, skip_exist, wait_sec):
        """ファイルを一括ダウンロード"""
        success, fail, skipped = 0, 0, 0

        for i, row in enumerate(targets, 1):
            if not self.is_running:
                break

            name = row.get('販売名', '')
            url = row.get('PDF_URL', '')
            doc_id = extract_doc_id_from_url(url)

            if not doc_id:
                continue

            # 販売名を10文字程度に制限し、ファイル名に使用できない文字を除去
            product_name_short = name[:10]
            # ファイル名に使用できない文字を置換
            invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
            for char in invalid_chars:
                product_name_short = product_name_short.replace(char, '_')

            # ファイル名を「doc_id_販売名.pdf」の形式にする
            if product_name_short:
                filename = f"{doc_id}_{product_name_short}{PDF_EXTENSION}"
            else:
                filename = f"{doc_id}{PDF_EXTENSION}"

            save_path = os.path.join(pdf_dir, filename)

            # ログ表示用には30文字に制限
            display_name = name[:30]
            self.log(f"[{i}/{len(targets)}] {display_name}...")

            if skip_exist and os.path.exists(save_path):
                self.log("  → スキップ(既存)")
                skipped += 1
            else:
                ok, err = download_file(url, save_path)
                if ok:
                    self.log("  → 完了")
                    success += 1
                else:
                    self.log(f"  → 失敗: {err}")
                    fail += 1
                time.sleep(wait_sec)

        return success, fail, skipped

    def _handle_download_completion(self, success, fail, skipped):
        """ダウンロード完了時の処理"""
        if not self.is_running:
            self.log("")
            self.log("=== 中断されました ===")
            self.log(f"成功: {success} / スキップ: {skipped} / 失敗: {fail}")
        else:
            self.log("")
            self.log("=== 全処理完了 ===")
            messagebox.showinfo("完了", f"ダウンロードが完了しました。\n\n成功: {success}\nスキップ: {skipped}\n失敗: {fail}")

    def download_thread(self, csv_file, out_dir, use_filter, hosp_csv, skip_exist, wait_sec):
        """ダウンロード処理のメインスレッド"""
        try:
            pdf_dir = os.path.join(out_dir, "pdf")
            os.makedirs(pdf_dir, exist_ok=True)

            target_approvals, target_certs = self._load_filter_lists(use_filter, hosp_csv)
            targets = self._load_and_filter_targets(csv_file, use_filter, target_approvals, target_certs)

            self.log(f"--- ダウンロード開始 (対象: {len(targets)}件) ---")

            success, fail, skipped = self._download_files(targets, pdf_dir, skip_exist, wait_sec)
            self._handle_download_completion(success, fail, skipped)

        except Exception as e:
            self.log(f"エラー: {e}")
            messagebox.showerror("エラー", str(e))
        finally:
            self.is_running = False
            self.set_running_state(False)


# ==========================================
# Step3: 設定タブ (SettingsTab)
# ==========================================
class SettingsTab(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, padding="10")
        self.app = app

        # 説明
        info_frame = ttk.LabelFrame(self, text="設定について", padding="10")
        info_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(info_frame, text="アプリケーション全体の動作設定を変更できます。", 
                 wraplength=650).pack(anchor=tk.W)
        ttk.Label(info_frame, text="設定は保存され、次回起動時にも反映されます。", 
                 wraplength=650, foreground="gray").pack(anchor=tk.W)

        # 基本設定
        basic_frame = ttk.LabelFrame(self, text="基本設定", padding="10")
        basic_frame.pack(fill=tk.X, pady=(0, 10))
        
        # デフォルト出力ディレクトリ
        dir_frame = ttk.Frame(basic_frame)
        dir_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(dir_frame, text="デフォルト出力ディレクトリ:", width=25).pack(side=tk.LEFT)
        self.output_dir_var = tk.StringVar(value=app_settings.default_output_dir)
        ttk.Entry(dir_frame, textvariable=self.output_dir_var, width=40).pack(side=tk.LEFT, padx=(5, 5))
        ttk.Button(dir_frame, text="参照...", command=self.browse_output_dir).pack(side=tk.LEFT)
        
        # ダウンロード待機時間
        wait_frame = ttk.Frame(basic_frame)
        wait_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(wait_frame, text="ダウンロード待機時間 (秒):", width=25).pack(side=tk.LEFT)
        self.wait_time_var = tk.StringVar(value=str(app_settings.default_wait_time))
        ttk.Entry(wait_frame, textvariable=self.wait_time_var, width=10).pack(side=tk.LEFT, padx=(5, 5))
        ttk.Label(wait_frame, text="※PDFダウンロード間隔", foreground="gray").pack(side=tk.LEFT)
        
        # スクレイピング待機時間
        scrape_wait_frame = ttk.Frame(basic_frame)
        scrape_wait_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(scrape_wait_frame, text="一覧取得待機時間 (秒):", width=25).pack(side=tk.LEFT)
        self.scrape_wait_var = tk.StringVar(value=str(app_settings.scrape_wait_time))
        ttk.Entry(scrape_wait_frame, textvariable=self.scrape_wait_var, width=10).pack(side=tk.LEFT, padx=(5, 5))
        ttk.Label(scrape_wait_frame, text="※日付間の待機時間", foreground="gray").pack(side=tk.LEFT)

        # 詳細設定（上級者向け）
        advanced_frame = ttk.LabelFrame(self, text="詳細設定（上級者向け）", padding="10")
        advanced_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(advanced_frame, text="⚠ 以下の設定を変更すると、正常に動作しなくなる可能性があります", 
                 foreground="red", wraplength=650).pack(anchor=tk.W, pady=(0, 5))
        
        # PMDA一覧ページURL
        url_frame = ttk.Frame(advanced_frame)
        url_frame.pack(fill=tk.X, pady=(5, 5))
        ttk.Label(url_frame, text="PMDA一覧ページURL:", width=25).pack(side=tk.LEFT)
        self.base_url_var = tk.StringVar(value=app_settings.base_url)
        ttk.Entry(url_frame, textvariable=self.base_url_var, width=50).pack(side=tk.LEFT, padx=(5, 0))
        
        # PMDA詳細ベースURL
        detail_url_frame = ttk.Frame(advanced_frame)
        detail_url_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(detail_url_frame, text="PMDA詳細ベースURL:", width=25).pack(side=tk.LEFT)
        self.detail_url_var = tk.StringVar(value=app_settings.detail_base_url)
        ttk.Entry(detail_url_frame, textvariable=self.detail_url_var, width=50).pack(side=tk.LEFT, padx=(5, 0))

        # ボタン
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(btn_frame, text="設定を保存", command=self.save_settings).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="デフォルトに戻す", command=self.reset_settings).pack(side=tk.LEFT, padx=(0, 10))
        self.status_label = ttk.Label(btn_frame, text="", foreground="green")
        self.status_label.pack(side=tk.LEFT, padx=(10, 0))

    def browse_output_dir(self):
        d = filedialog.askdirectory(initialdir=self.output_dir_var.get())
        if d:
            self.output_dir_var.set(d)

    def save_settings(self):
        try:
            # 入力値の検証
            wait_time = float(self.wait_time_var.get())
            scrape_wait = float(self.scrape_wait_var.get())
            
            if wait_time < MIN_WAIT_TIME or wait_time > MAX_WAIT_TIME:
                messagebox.showerror("エラー", f"ダウンロード待機時間は{MIN_WAIT_TIME}～{MAX_WAIT_TIME}秒の範囲で指定してください")
                return

            if scrape_wait < MIN_WAIT_TIME or scrape_wait > MAX_WAIT_TIME:
                messagebox.showerror("エラー", f"一覧取得待機時間は{MIN_WAIT_TIME}～{MAX_WAIT_TIME}秒の範囲で指定してください")
                return
            
            # 設定を更新
            app_settings.default_output_dir = self.output_dir_var.get()
            app_settings.default_wait_time = wait_time
            app_settings.scrape_wait_time = scrape_wait
            app_settings.base_url = self.base_url_var.get()
            app_settings.detail_base_url = self.detail_url_var.get()
            
            # ファイルに保存
            app_settings.save()

            # 各タブのUIを即座に更新
            self.app.update_ui_from_settings()

            self.status_label.configure(text="✓ 設定を保存しました")
            self.after(3000, lambda: self.status_label.configure(text=""))

            messagebox.showinfo("保存完了", "設定を保存しました。\n新しい設定がすべてのタブに反映されました。")
            
        except ValueError:
            messagebox.showerror("エラー", "待機時間は数値で入力してください")

    def reset_settings(self):
        if messagebox.askyesno("確認", "すべての設定をデフォルト値に戻しますか？"):
            app_settings.reset()
            app_settings.save()

            # 設定タブのUIを更新
            self.output_dir_var.set(app_settings.default_output_dir)
            self.wait_time_var.set(str(app_settings.default_wait_time))
            self.scrape_wait_var.set(str(app_settings.scrape_wait_time))
            self.base_url_var.set(app_settings.base_url)
            self.detail_url_var.set(app_settings.detail_base_url)

            # 各タブのUIを即座に更新
            self.app.update_ui_from_settings()

            self.status_label.configure(text="✓ デフォルト値に戻しました")
            self.after(3000, lambda: self.status_label.configure(text=""))

            messagebox.showinfo("リセット完了", "設定をデフォルト値に戻しました。")


# ==========================================
# メインアプリケーション
# ==========================================
class PMDAToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PMDA添付文書DLツール")

        # 設定を読み込み
        app_settings.load()

        # ウィンドウサイズと位置を復元
        self.restore_window_geometry()

        # ウィンドウ閉じる時のイベントハンドラを設定
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # タブスタイルをカスタマイズ（縁取りを追加して見やすく）
        style = ttk.Style()
        style.theme_use('default')  # デフォルトテーマを使用

        # タブの枠線と色を設定
        style.configure('TNotebook', borderwidth=2, relief='solid')
        style.configure('TNotebook.Tab',
                       padding=[20, 8],  # タブの内側の余白を増やす
                       borderwidth=2,     # 枠線の太さ
                       relief='raised',   # 立体的な枠線
                       background='#e0e0e0')  # 背景色（グレー）

        # 選択されているタブの色
        style.map('TNotebook.Tab',
                 background=[('selected', '#ffffff')],  # 選択時は白
                 relief=[('selected', 'sunken')],       # 選択時は凹んだ感じ
                 borderwidth=[('selected', 2)])

        # タブ設定
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.scrape_tab = ScrapeTab(self.notebook, self)
        self.download_tab = DownloadTab(self.notebook)
        self.settings_tab = SettingsTab(self.notebook, self)

        self.notebook.add(self.scrape_tab, text="Step 1: 一覧取得")
        self.notebook.add(self.download_tab, text="Step 2: PDFダウンロード")
        self.notebook.add(self.settings_tab, text="設定")

    def set_download_csv_path(self, path):
        """一覧取得完了時にダウンロードタブへパスを渡してタブを切り替える"""
        self.download_tab.csv_path_var.set(path)
        self.notebook.select(self.download_tab)

    def update_ui_from_settings(self):
        """設定変更を各タブのUIに即座に反映"""
        # 一覧取得タブの出力先を更新
        self.scrape_tab.out_dir_var.set(app_settings.default_output_dir)

        # ダウンロードタブの出力先を更新
        self.download_tab.out_dir_var.set(os.path.join(app_settings.default_output_dir, "downloads"))

    def restore_window_geometry(self):
        """保存されたウィンドウサイズと位置を復元"""
        width = app_settings.window_width
        height = app_settings.window_height
        x = app_settings.window_x
        y = app_settings.window_y

        if x is not None and y is not None:
            # 位置も復元
            self.root.geometry(f"{width}x{height}+{x}+{y}")
        else:
            # 位置が保存されていない場合は、サイズのみ設定
            self.root.geometry(f"{width}x{height}")

    def save_window_geometry(self):
        """現在のウィンドウサイズと位置を保存"""
        # ウィンドウの状態を取得
        geometry = self.root.geometry()  # 例: "700x680+100+50"

        # サイズと位置を分解
        size_pos = geometry.split('+')
        if len(size_pos) >= 1:
            size = size_pos[0].split('x')
            if len(size) == 2:
                app_settings.window_width = int(size[0])
                app_settings.window_height = int(size[1])

        if len(size_pos) >= 3:
            app_settings.window_x = int(size_pos[1])
            app_settings.window_y = int(size_pos[2])

        app_settings.save()

    def on_closing(self):
        """ウィンドウを閉じる時の処理"""
        # ウィンドウサイズと位置を保存
        self.save_window_geometry()

        # ウィンドウを閉じる
        self.root.destroy()


def main():
    root = tk.Tk()
    app = PMDAToolApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()
