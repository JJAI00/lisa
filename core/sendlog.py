import os
import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta

def resource_path(relative_path):
    import sys, os
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.abspath('.')
    return os.path.join(base_path, relative_path)

# 구글 시트 접속 정보
SCRIPT_DIR     = os.path.dirname(os.path.abspath(__file__))
KEY_FILE       = resource_path('google_sheet_key/renewal-bot-463605-b8f41de8fbdb.json')
SCOPES         = ['https://www.googleapis.com/auth/spreadsheets']
LOG_SHEET_ID   = '1llolEKlleT6Cve_M5Qt_CO_JDxQczIKBMqL8fj5vfSk'
LOG_SHEET_NAME = 'Sheet1'

def show_logs(parent):
    """
    별도 창을 띄워 기간별로 로그를 표시합니다.
    - 오늘: 자정 이후 기록
    - 일주일/한달/세달: 각각 7/30/90일 전부터
    - 모두 최신순(내림차순) 정렬
    """
    try:
        creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
        gc    = gspread.authorize(creds)
        sh    = gc.open_by_key(LOG_SHEET_ID)
        ws    = sh.worksheet(LOG_SHEET_NAME)
        records = ws.get_all_records()
        log_df = pd.DataFrame(records)
        if '일시' in log_df.columns:
            log_df['일시'] = pd.to_datetime(
                log_df['일시'], format='%Y-%m-%d %H:%M', errors='coerce'
            )
            log_df.sort_values('일시', ascending=False, inplace=True)
        log_df.columns = log_df.columns.str.strip()
    except Exception as e:
        messagebox.showerror("로그 오류", f"로그를 불러오는 데 실패했습니다:\n{e}")
        return

    win = tk.Toplevel(parent)
    win.title("로그 보기")

    toolbar = tk.Frame(win)
    toolbar.pack(fill='x', pady=5)

    def refresh(days):
        now = datetime.now()
        if days == 0:
            # 오늘 자정 이후
            cutoff = now.replace(hour=0, minute=0, second=0, microsecond=0)
        else:
            cutoff = now - timedelta(days=days)
        if '일시' in log_df.columns:
            filtered = log_df[log_df['일시'] >= cutoff].copy()
            filtered.sort_values('일시', ascending=False, inplace=True)
        else:
            filtered = log_df.copy()

        tree.delete(*tree.get_children())
        for _, row in filtered.iterrows():
            tree.insert('', 'end', values=list(row.values))

    tk.Button(toolbar, text="오늘",   command=lambda: refresh(0)).pack(side='left', padx=5)
    tk.Button(toolbar, text="일주일", command=lambda: refresh(7)).pack(side='left', padx=5)
    tk.Button(toolbar, text="한달",   command=lambda: refresh(30)).pack(side='left', padx=5)
    tk.Button(toolbar, text="세달",   command=lambda: refresh(90)).pack(side='left', padx=5)

    cols = list(log_df.columns)
    tree = ttk.Treeview(win, columns=cols, show='headings')
    vsb  = ttk.Scrollbar(win, orient='vertical',   command=tree.yview)
    hsb  = ttk.Scrollbar(win, orient='horizontal', command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tree.pack(side='left', fill='both', expand=True)
    vsb.pack(side='right', fill='y')
    hsb.pack(side='bottom', fill='x')

    for col in cols:
        tree.heading(col, text=col)
        tree.column(col, width=120, anchor='center')

    # 초기 표시: 오늘
    refresh(0)
