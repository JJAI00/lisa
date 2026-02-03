# renewal_email_multi.py

import sys, os
import re
import pandas as pd
from datetime import datetime
import win32com.client as win32
from email_module.renewal_email import append_log, TEMPLATE_DIR

def safe_parse_date(date_value):
    """안전한 날짜 파싱 함수"""
    if pd.isna(date_value) or date_value is None or date_value == '':
        return None
    
    if isinstance(date_value, str):
        date_value = date_value.strip()
        if not date_value:
            return None
    
    try:
        if isinstance(date_value, str):
            return datetime.strptime(date_value, '%Y-%m-%d')
        elif hasattr(date_value, 'strftime'):
            return date_value
        else:
            return None
    except (ValueError, TypeError):
        return None

def resource_path(relative_path):
    import sys, os
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.abspath('.')
    return os.path.join(base_path, relative_path)

# TEMPLATE_DIR = resource_path('email_temp')
# 기타 파일 접근도 동일하게 resource_path 사용

def send_multi_emails(rows, user):
    """
    동일 거래처의 여러 건을 하나의 메일로 발송합니다.
    - rows: 리스트 형태의 dict(row)들
    - user: 실행자(로그 기록용)
    """
    if not rows:
        return

    # 1) HTML 템플릿 선택 (첫 번째 행의 대분류 기준)
    first = rows[0]
    tpl = os.path.join(
        TEMPLATE_DIR,
        f"email_template_{first.get('대분류','').strip()}.html"
    )
    if not os.path.exists(tpl):
        tpl = os.path.join(TEMPLATE_DIR, 'email_template.html')

    # 2) 원본 템플릿 읽기
    try:
        with open(tpl, encoding='utf-8') as f:
            body_raw = f.read()
    except FileNotFoundError:
        body_raw = ""

    # 3) 템플릿 내 첫 번째 placeholder 행 제거 (테이블 첫줄 제외)
    #    {customer}가 포함된 <tr>...</tr> 블록을 삭제
    body_clean = re.sub(
        r"<tr>.*?\{customer\}.*?</tr>\s*",
        "",
        body_raw,
        flags=re.DOTALL
    )

    # 4) {name}, {customer} 치환
    data = {
        'name':     first.get('담당자명', ''),
        'customer': first.get('고객사명', '')
    }
    try:
        body = body_clean.format(**data)
    except Exception:
        body = body_clean

    # 5) 동적 테이블 행 생성 (고객사명, 제품, 수량, 만료일만)
    table_rows = ""
    for row in rows:
        cust   = row['고객사명']
        prod   = row['제품']
        qty    = row['수량']
        
        # 안전한 날짜 처리
        expiry_date = safe_parse_date(row['만료일'])
        if expiry_date:
            expiry = expiry_date.strftime('%Y-%m-%d')
        else:
            expiry = str(row['만료일']) if row['만료일'] else ''
        
        table_rows += (
            f"<tr>"
            f"<td>{cust}</td>"
            f"<td>{prod}</td>"
            f"<td>{qty}</td>"
            f"<td>{expiry}</td>"
            f"</tr>\n"
        )

    # 6) </table> 직전에 동적 행 삽입
    if "</table>" in body:
        body_multi = body.replace("</table>", f"{table_rows}</table>")
    else:
        body_multi = body + table_rows

    # 7) Outlook 메일 생성 및 설정 (메인 스레드에서 직접)
    try:
        outlook = win32.Dispatch('Outlook.Application')
        print("아웃룩 연결 성공")
        
        mail = outlook.CreateItem(0)
        mail.HTMLBody = body_multi

        # 받는 사람(To) 자동 설정 생략 — 테스트나 수동입력을 위해 빈칸으로 둡니다.
        # tos = ";".join({r.get('이메일','') for r in rows if r.get('이메일')})
        # mail.To = tos

        # 제목: 첫 번째 제품명 + X건
        extra = len(rows) - 1
        if extra > 0:
            mail.Subject = f"[큐브렉스] {first['제품']} 외 {extra}건 라이센스 리뉴얼 안내"
        else:
            mail.Subject = f"[큐브렉스] {first['제품']} 라이센스 리뉴얼 안내"

        mail.Display()  # 미리보기
        
    except Exception as e:
        error_msg = f"아웃룩 연결 실패: {str(e)}\n\n아웃룩을 실행한 후 다시 시도해주세요."
        import tkinter.messagebox as messagebox
        messagebox.showerror("아웃룩 연결 오류", error_msg)
        return

    # 8) 각 행별로 로그 기록
    for row in rows:
        append_log(row, action='메일 작성', user=user)
