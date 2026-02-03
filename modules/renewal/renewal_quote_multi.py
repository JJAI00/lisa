import sys, os
from datetime import datetime
import win32com.client as win32
from tkinter import messagebox
import re
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

# renewal_quote.py에서 import한 경로 상수도 resource_path로 접근하도록 보장
# 예시: template_path = resource_path('quote/quote_origin_01.xlsx')
# 기타 파일 접근도 동일하게 resource_path 사용

# renewal_quote.py의 상수 및 함수 가져오기
from .renewal_quote import (
    SCRIPT_DIR, KEY_FILE, SCOPES,
    LOG_SHEET_ID, LOG_SHEET_NAME,
    QUOTE_TEMPLATE_PATH, QUOTE_EMAIL_DIR,
    append_log
)


def create_and_send_quotes(rows, user=None):
    """
    다건 견적서 작성 및 전송 (Multi Quote)
    rows: list of dict 또는 pandas.Series for 같은 고객사 제품들
    user: 작성자 이름
    """
    # Foundry 대분류 항목 제외
    filtered = [r for r in rows if r.get('대분류') != 'Foundry']
    if not filtered:
        messagebox.showinfo("알림", "모두 Foundry 제품이므로 견적서 작성 대상이 없습니다.")
        return
    rows = filtered

    # 템플릿 파일 확인
    if not os.path.exists(QUOTE_TEMPLATE_PATH):
        messagebox.showerror(
            "템플릿 파일 오류",
            f"견적서 템플릿 파일을 찾을 수 없습니다:\n{QUOTE_TEMPLATE_PATH}"
        )
        return

    # Excel 애플리케이션 실행
    excel = win32.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None

    try:
        # 템플릿 불러오기
        wb = excel.Workbooks.Open(QUOTE_TEMPLATE_PATH)
        ws = wb.Worksheets(1)

        # 공통 헤더 정보 입력
        today_str = datetime.now().strftime('%Y-%m-%d')
        ws.Range("E3").Value = today_str
        customer = rows[0].get('고객사명', '')
        ws.Range("C4").Value = customer

        # 담당자, 연락처, 이메일 (첫 행 기준)
        ws.Range("E6").Value = rows[0].get('담당자명', '')
        ws.Range("E7").Value = rows[0].get('연락처', '')
        ws.Range("E8").Value = rows[0].get('이메일', '')

        # 다건 제품 정보 입력 (제품명, 수량, 단가, 원가, 만료일)
        start_row = 14
        for idx, row in enumerate(rows):
            r = start_row + idx
            ws.Range(f"D{r}").Value = row.get('제품', '')
            ws.Range(f"L{r}").Value = row.get('수량', '')
            ws.Range(f"M{r}").Value = row.get('판매 단가', '')
            ws.Range(f"T{r}").Value = row.get('원가', '')
            # 만료일 안전 처리
            expiry = row.get('만료일')
            if pd.isna(expiry):
                expiry_str = ''
            else:
                try:
                    if hasattr(expiry, 'strftime'):
                        expiry_str = expiry.strftime('%Y-%m-%d')
                    else:
                        expiry_str = str(expiry)
                except:
                    expiry_str = str(expiry)
            ws.Range(f"Q{r}").Value = expiry_str

        # 파일명 포맷 설정: 고객사명_제품명 외_날짜
        if len(rows) == 1:
            prod_str = rows[0].get('제품', '')
        else:
            prod_str = '제품명 외'
        base_name = f"{customer}_{prod_str}_{today_str}"

        # 저장 경로 준비
        out_dir = os.path.join(resource_root(), 'quote')
        os.makedirs(out_dir, exist_ok=True)
        xlsx_path = os.path.join(out_dir, base_name + '.xlsx')
        pdf_path  = os.path.join(out_dir, base_name + '.pdf')

        # 파일 저장
        wb.SaveAs(xlsx_path)
        ws.ExportAsFixedFormat(0, pdf_path)

    except Exception as e:
        messagebox.showerror("Excel 처리 오류", str(e))
        return

    finally:
        if wb:
            wb.Close(False)
        excel.Quit()

    # Outlook 메일 작성 및 템플릿 로드
    try:
        # 아웃룩 연결 시도 (메인 스레드에서 직접)
        outlook = win32.Dispatch('Outlook.Application')
        print("아웃룩 연결 성공")
        
        # Outlook 앱
        mail = outlook.CreateItem(0)
        # 받는 사람 설정: 중복 제거
        tos = ";".join({r.get('이메일','') for r in rows if r.get('이메일')})
        mail.To = tos

        # HTML 템플릿 선택
        cat = rows[0].get('대분류','').strip()
        tmpl = {
            'TeamViewer': 'email_template_TeamViewer.html',
            'Foundry':    'email_template_Foundry.html',
            'Chaos':      'email_template_Chaosgroup.html'
        }.get(cat, 'email_template.html')
        tmpl_path = os.path.join(QUOTE_EMAIL_DIR, tmpl)

        # 템플릿 읽기
        try:
            with open(tmpl_path, encoding='utf-8') as f:
                body_raw = f.read()
        except FileNotFoundError:
            body_raw = ''

        # 기존 단일 placeholder 행 제거
        body_clean = re.sub(
            r"<tr>.*?\{customer\}.*?</tr>\s*",
            "",
            body_raw,
            flags=re.DOTALL
        )

        # 공통 치환
        data_header = {
            'name':     rows[0].get('담당자명',''),
            'customer': customer
        }
        try:
            body_header = body_clean.format(**data_header)
        except Exception:
            body_header = body_clean

        # 동적 테이블 행 생성
        table_rows = ''
        for r in rows:
            # 만료일 안전 처리
            expiry_val = r.get('만료일')
            if pd.isna(expiry_val):
                expiry_str = 'N/A'
            else:
                try:
                    if isinstance(expiry_val, str):
                        expiry_str = expiry_val
                    elif hasattr(expiry_val, 'strftime'):
                        expiry_str = expiry_val.strftime('%Y-%m-%d')
                    else:
                        expiry_str = str(expiry_val)
                except:
                    expiry_str = str(expiry_val)
            
            table_rows += (
                f"<tr>"
                f"<td>{r.get('고객사명','')}</td>"
                f"<td>{r.get('제품','')}</td>"
                f"<td>{r.get('수량','')}</td>"
                f"<td>{expiry_str}</td>"
                f"</tr>\n"
            )

        # 동적 행 삽입
        if '</table>' in body_header:
            body_multi = body_header.replace('</table>', f"{table_rows}</table>")
        else:
            body_multi = body_header + table_rows

        # 메일 본문 설정
        mail.HTMLBody = body_multi

        # 제목 설정
        first_prod = rows[0].get('제품','')
        extra = len(rows) - 1
        if extra > 0:
            mail.Subject = f"[큐브렉스] {first_prod} 외 {extra}건 견적서 안내 ({today_str})"
        else:
            mail.Subject = f"[큐브렉스] {first_prod} 견적서 안내 ({today_str})"

        # 첨부
        mail.Attachments.Add(pdf_path)
        mail.Display()

    except Exception as e:
        error_msg = f"아웃룩 연결 실패: {str(e)}\n\n아웃룩을 실행한 후 다시 시도해주세요."
        messagebox.showerror("아웃룩 연결 오류", error_msg)
        return

    # 로그 기록
    for row in rows:
        append_log(row, action='메일+견적서', user=user)

# 이전 버전 호출 호환성
create_and_send_quote = create_and_send_quotes
