# renewal_quote.py

import os
import sys
from datetime import datetime
import win32com.client as win32
import gspread
from google.oauth2.service_account import Credentials
from tkinter import messagebox
import pandas as pd

def resource_path(relative_path):
    import sys, os
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.abspath('.')
    return os.path.join(base_path, relative_path)

def resource_root():
    import sys, os
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.abspath('.')

# —————— 설정 값 ——————
SCRIPT_DIR       = os.path.dirname(os.path.abspath(__file__))
KEY_FILE         = resource_path('google_sheet_key/renewal-bot-463605-b8f41de8fbdb.json')
SCOPES           = ['https://www.googleapis.com/auth/spreadsheets']
LOG_SHEET_ID     = '1llolEKlleT6Cve_M5Qt_CO_JDxQczIKBMqL8fj5vfSk'
LOG_SHEET_NAME   = 'Sheet1'

QUOTE_TEMPLATE_PATH = resource_path('quote/quote_origin_01.xlsx')
QUOTE_EMAIL_DIR     = resource_path('quote_email_temp')

def append_log(row, action='메일+견적서', user=None):
    import socket
    import pandas as pd
    
    executor = socket.gethostname()
    
    # 만료일 안전 처리
    expiry_date = row['만료일']
    if pd.isna(expiry_date):
        expiry_str = 'N/A'
    else:
        try:
            expiry_str = expiry_date.strftime('%Y-%m-%d')
        except:
            expiry_str = str(expiry_date)
    
    log_data = [
        row['고객사명'],
        row.get('담당자명', ''),
        row.get('연락처', ''),
        row.get('이메일', ''),
        row['제품'],
        row['수량'],
        expiry_str,
        action,
        executor,
        datetime.now().strftime('%Y-%m-%d %H:%M')
    ]
    creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(LOG_SHEET_ID)
    try:
        ws = sh.worksheet(LOG_SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=LOG_SHEET_NAME, rows='1000', cols='20')
        ws.append_row(
            ['고객사명','담당자명','연락처','이메일','제품','수량','만료일','액션','실행자','일시','사용자'],
            value_input_option='USER_ENTERED'
        )
    ws.append_row(log_data + [user], value_input_option='USER_ENTERED')

def create_and_send_quote(row, user=None):
    if not os.path.exists(QUOTE_TEMPLATE_PATH):
        messagebox.showerror(
            "템플릿 파일 오류",
            f"견적서 템플릿(.xlsx) 파일을 찾을 수 없습니다:\n{QUOTE_TEMPLATE_PATH}"
        )
        return

    excel = win32.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = None
    try:
        wb = excel.Workbooks.Open(QUOTE_TEMPLATE_PATH, ReadOnly=True)
        ws = wb.Worksheets(1)

        today_str = datetime.now().strftime('%Y-%m-%d')
        ws.Range("E3").Value  = today_str
        ws.Range("C4").Value  = row['고객사명']
        ws.Range("E6").Value  = row.get('담당자명','')
        ws.Range("E7").Value  = row.get('연락처','')
        ws.Range("E8").Value  = row.get('이메일','')

        ws.Range("D14").Value = row['제품']
        ws.Range("L14").Value = row['수량']
        ws.Range("M14").Value = row.get('판매 단가', '')  # 단가
        ws.Range("T14").Value = row.get('판매 원가', '')  # 원가
        
        # 만료일 안전 처리
        expiry_date = row['만료일']
        if pd.isna(expiry_date):
            expiry_str = ''
        else:
            try:
                expiry_str = expiry_date.strftime('%Y-%m-%d')
            except:
                expiry_str = str(expiry_date)
        ws.Range("Q14").Value = expiry_str

        base = f"{row['고객사명']}_{row['제품']}_{today_str}"
        out_dir = os.path.join(resource_root(), 'quote')
        os.makedirs(out_dir, exist_ok=True)
        xlsx_path = os.path.join(out_dir, base + '.xlsx')
        pdf_path  = os.path.join(out_dir, base + '.pdf')

        wb.SaveAs(xlsx_path)
        ws.ExportAsFixedFormat(0, pdf_path)

    except Exception as e:
        messagebox.showerror("Excel 처리 오류", str(e))

    finally:
        if wb:
            wb.Close(False)
        excel.Quit()

    try:
        # 아웃룩 연결 시도 (메인 스레드에서 직접)
        outlook = win32.Dispatch('Outlook.Application')
        print("아웃룩 연결 성공")
        
        mail = outlook.CreateItem(0)
        mail.To = row.get('이메일','')

        cat = row.get('대분류','').strip()
        tmpl = {
            'TeamViewer': 'email_template_TeamViewer.html',
            'Foundry':    'email_template_Foundry.html',
            'Chaos':      'email_template_Chaosgroup.html'
        }.get(cat, 'email_template.html')
        tmpl_path = os.path.join(QUOTE_EMAIL_DIR, tmpl)
        body = ''
        if os.path.exists(tmpl_path):
            with open(tmpl_path, encoding='utf-8') as f:
                body = f.read()

        # 만료일 안전 처리
        expiry_date = row['만료일']
        if pd.isna(expiry_date):
            expiry_str = 'N/A'
        else:
            try:
                if isinstance(expiry_date, str):
                    expiry_str = expiry_date
                else:
                    expiry_str = expiry_date.strftime('%Y-%m-%d')
            except:
                expiry_str = str(expiry_date)
        
        data = {
            'customer': row['고객사명'],
            'product':  row['제품'],
            'quantity': row['수량'],
            'expiry':   expiry_str,
            'name':     row.get('담당자명','')
        }
        data['qty'] = data['quantity']

        try:
            mail.HTMLBody = body.format(**data)
        except Exception:
            mail.HTMLBody = body

        mail.Subject = f"[큐브렉스] {row['제품']} 견적서 및 리뉴얼 안내 ({today_str})"
        mail.Attachments.Add(pdf_path)
        mail.Display()
        
    except Exception as e:
        error_msg = f"아웃룩 연결 실패: {str(e)}\n\n아웃룩을 실행한 후 다시 시도해주세요."
        messagebox.showerror("아웃룩 연결 오류", error_msg)
        return

    append_log(row, action='메일+견적서', user=user)
