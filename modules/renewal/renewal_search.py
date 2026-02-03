import os
import pandas as pd
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
import gspread
from config import KEY_FILE, SCOPES, SHEET_IDS, WORKSHEET_NAMES

# 하위 호환성을 위한 별칭
JSON_KEYFILE = KEY_FILE

# Google Sheet IDs (config에서 가져옴)
RENEWAL_LIST_SHEET_ID  = SHEET_IDS['renewal_list']
CUSTOMER_LIST_SHEET_ID = SHEET_IDS['customer_list']
RENEWAL_SHEET_NAME     = WORKSHEET_NAMES['renewal_list']
CUSTOMER_SHEET_NAME    = WORKSHEET_NAMES['customer_list']

def _authorize():
    creds = Credentials.from_service_account_file(JSON_KEYFILE, scopes=SCOPES)
    return gspread.authorize(creds)

def get_filtered_data_range(start_date, end_date):
    """
    renewal_list 시트에서 만료일이 start_date~end_date 범위인 데이터 가져오기
    그리고 customer_list 시트의 고객사 정보를 결합
    """
    gc = _authorize()
    # ── renewal_list 불러오기 ──
    sh = gc.open_by_key(RENEWAL_LIST_SHEET_ID)
    ws = sh.worksheet(RENEWAL_SHEET_NAME)
    df = pd.DataFrame(ws.get_all_records())
    # 만료일을 datetime64[ns]로 변환
    df['만료일'] = pd.to_datetime(df['만료일'], errors='coerce')
    # start_date, end_date를 pd.Timestamp로 변환
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)
    df = df[(df['만료일'] >= start_date) & (df['만료일'] <= end_date)]

    # ── customer_list 불러오기 (헤더 중복 회피) ──
    sh2 = gc.open_by_key(CUSTOMER_LIST_SHEET_ID)
    ws2 = sh2.worksheet(CUSTOMER_SHEET_NAME)
    all_values = ws2.get_all_values()
    if not all_values:
        contacts = pd.DataFrame()
    else:
        # 1행을 헤더로, 빈 문자열 헤더는 제외
        raw_header = [h.strip() for h in all_values[0]]
        valid_indices = [i for i, h in enumerate(raw_header) if h]
        headers = [raw_header[i] for i in valid_indices]
        rows = []
        for row in all_values[1:]:
            # 각 행에서 빈 헤더 열은 건너뛴 뒤 값 수집
            rows.append([row[i] for i in valid_indices])
        contacts = pd.DataFrame(rows, columns=headers)

    # ── 고객사명 기준으로 병합 ──
    # '고객사명' 컬럼 이름이 정확히 일치해야 합니다.
    df = df.merge(contacts, on='고객사명', how='left')
    return df

# ── 기간 계산 함수들 ──

def period_next_4weeks(base):
    monday = base + timedelta(days=(0 - base.weekday()) % 7, weeks=4)
    sunday = monday + timedelta(days=6)
    return monday, sunday

def period_next_month(base):
    year  = base.year + (base.month // 12)
    month = (base.month % 12) + 1
    start = base.replace(year=year, month=month, day=1)
    nm    = start.replace(day=28) + timedelta(days=4)
    end   = nm - timedelta(days=nm.day)
    return start, end

def period_one_year(base):
    return base, base.replace(year=base.year+1)

def period_past_year(base):
    return base.replace(year=base.year-1), base

# ── 필터 헬퍼 ──

def get_categories():
    df = get_filtered_data_range(datetime.today().date(), datetime.today().date())
    return sorted(df['대분류'].dropna().unique().tolist())

def filter_by_category(df, category):
    if category and category != '전체':
        return df[df['대분류'] == category]
    return df

def filter_by_name(df, name):
    if name:
        return df[df['고객사명'].str.contains(name, na=False)]
    return df
