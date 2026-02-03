# renewal_email.py

import sys, os
import socket
from datetime import datetime
import win32com.client as win32
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

def resource_path(relative_path):
    import sys, os
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.abspath('.')
    return os.path.join(base_path, relative_path)

# 서비스 계정 키 및 권한 스코프
KEY_FILE = resource_path('google_sheet_key/renewal-bot-463605-b8f41de8fbdb.json')
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
LOG_SHEET_ID = '1llolEKlleT6Cve_M5Qt_CO_JDxQczIKBMqL8fj5vfSk'
LOG_SHEET_NAME = 'Sheet1'

# 이메일 템플릿 디렉토리
TEMPLATE_DIR = resource_path('email_temp')

def check_outlook_connection():
    """아웃룩 연결 상태를 확인하는 함수"""
    try:
        outlook = win32.Dispatch('Outlook.Application')
        print("아웃룩 연결 상태: 정상")
        return True
    except Exception as e:
        print(f"아웃룩 연결 상태: 연결 실패 - {e}")
        return False

def append_log(row, action='메일 작성', user=None):
    executor = socket.gethostname()
    expiry = ''
    if '만료일' in row and isinstance(row['만료일'], pd.Timestamp) and not pd.isna(row['만료일']):
        expiry = row['만료일'].strftime('%Y-%m-%d')
    elif '만료일' in row and isinstance(row['만료일'], str) and row['만료일'].strip():
        expiry = row['만료일']
    log_data = [
        row['고객사명'],
        row.get('담당자명', ''),
        row.get('연락처', ''),
        row.get('이메일', ''),
        row['제품'],
        row['수량'],
        expiry,
        action,
        executor,
        datetime.now().strftime('%Y-%m-%d %H:%M')
    ]
    # 구글 스프레드시트 인증 및 시트 열기
    creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(LOG_SHEET_ID)
    try:
        ws = sh.worksheet(LOG_SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        # 로그 시트가 없으면 새로 생성하고 헤더 추가 (K열: 사용자)
        ws = sh.add_worksheet(title=LOG_SHEET_NAME, rows='1000', cols='20')
        ws.append_row(
            ['고객사명','담당자명','연락처','이메일','제품','수량','만료일','액션','실행자','일시','사용자'],
            value_input_option='USER_ENTERED'
        )
    # 실제 로그 데이터에 사용자 ID를 덧붙여 기록
    ws.append_row(log_data + [user], value_input_option='USER_ENTERED')

def send_emails(rows, user):
    """
    선택된 행에 대해 리뉴얼 안내 메일을 작성하고 로그를 남깁니다.
    """
    # 아웃룩 연결 시도 (메인 스레드에서 직접)
    try:
        outlook = win32.Dispatch('Outlook.Application')
        print("아웃룩 연결 성공")
        
    except Exception as e:
        error_msg = f"아웃룩 연결 실패: {str(e)}\n\n아웃룩을 실행한 후 다시 시도해주세요."
        import tkinter.messagebox as messagebox
        messagebox.showerror("아웃룩 연결 오류", error_msg)
        return
    
    for row in rows:
        try:
            mail = outlook.CreateItem(0)

            # 제품 대분류별 HTML 템플릿 선택
            tpl = os.path.join(
                TEMPLATE_DIR,
                f"email_template_{row.get('대분류','').strip()}.html"
            )
            if not os.path.exists(tpl):
                tpl = os.path.join(TEMPLATE_DIR, 'email_template.html')
            try:
                with open(tpl, encoding='utf-8') as f:
                    body = f.read()
            except FileNotFoundError:
                body = ''

            # 템플릿 포맷용 데이터 딕셔너리
            expiry = ''
            if '만료일' in row and isinstance(row['만료일'], pd.Timestamp) and not pd.isna(row['만료일']):
                expiry = row['만료일'].strftime('%Y-%m-%d')
            elif '만료일' in row and isinstance(row['만료일'], str) and row['만료일'].strip():
                expiry = row['만료일']
            data = {
                'customer': row['고객사명'],
                'product':  row['제품'],
                'quantity': row['수량'],
                'expiry':   expiry,
                'name':     row.get('담당자명','')
            }
            data['qty'] = data['quantity']

            # 본문 HTML 적용
            try:
                mail.HTMLBody = body.format(**data)
            except Exception:
                mail.HTMLBody = body

            mail.To      = row.get('이메일','')
            mail.Subject = f"[큐브렉스] {row['제품']} 라이센스 리뉴얼 안내"
            mail.Display()  # 미리보기

            # 로그에 사용자 정보와 함께 기록
            append_log(row, action='메일 작성', user=user)
            
        except Exception as e:
            error_msg = f"메일 작성 중 오류 발생: {str(e)}"
            print(error_msg)
            import tkinter.messagebox as messagebox
            messagebox.showerror("메일 작성 오류", error_msg)
            continue
