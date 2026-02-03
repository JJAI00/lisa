# utils/date_utils.py
"""날짜 관련 유틸리티 함수"""

import pandas as pd
from datetime import datetime, timedelta


def parse_korean_date(date_str):
    """
    한국어 날짜 형식을 파싱하여 datetime 객체로 변환
    
    지원 형식:
    - "2025. 8. 4 오전 9:25:07"
    - "2025. 8. 4"
    - "2025-08-01 17:25:23"
    - "2025-08-01"
    - 기타 pandas가 자동 인식하는 형식
    
    Args:
        date_str: 날짜 문자열
        
    Returns:
        datetime 객체 또는 None
    """
    if pd.isna(date_str) or date_str == '' or date_str is None:
        return None
    
    try:
        # 먼저 pandas 자동 파싱 시도
        return pd.to_datetime(date_str)
    except:
        pass
    
    try:
        # 한국어 형식 파싱 (예: "2025. 8. 4 오전 9:25:07")
        if isinstance(date_str, str):
            # "2025. 8. 4 오전 9:25:07" 형식 처리
            if '오전' in date_str or '오후' in date_str:
                # 오전/오후를 AM/PM으로 변환
                date_str_converted = date_str.replace('오전', 'AM').replace('오후', 'PM')
                return pd.to_datetime(date_str_converted, format='%Y. %m. %d %p %I:%M:%S')
            else:
                # "2025. 8. 4" 형식 처리
                if '. ' in date_str:
                    return pd.to_datetime(date_str, format='%Y. %m. %d')
    except:
        pass
    
    try:
        # "2025-08-01 17:25:23" 또는 "2025-08-01" 형식 처리
        if isinstance(date_str, str) and '-' in date_str:
            if ':' in date_str:
                return pd.to_datetime(date_str, format='%Y-%m-%d %H:%M:%S')
            else:
                return pd.to_datetime(date_str, format='%Y-%m-%d')
    except:
        pass
    
    # 모든 시도 실패
    return None


def get_period_dates(mode, base_date=None):
    """
    기간 모드에 따른 시작일과 종료일 계산
    
    Args:
        mode: 'this_month', 'next_month', 'one_year', 'past_year'
        base_date: 기준일 (기본값: 오늘)
        
    Returns:
        tuple: (시작일, 종료일)
    """
    if base_date is None:
        base_date = datetime.today()
    elif isinstance(base_date, str):
        base_date = pd.to_datetime(base_date)
    
    if mode == 'this_month':
        # 이번달
        start = base_date.replace(day=1)
        # 다음달 1일에서 하루 빼기
        if base_date.month == 12:
            end = base_date.replace(year=base_date.year + 1, month=1, day=1) - timedelta(days=1)
        else:
            end = base_date.replace(month=base_date.month + 1, day=1) - timedelta(days=1)
        return start, end
    
    elif mode == 'next_month':
        # 다음달
        if base_date.month == 12:
            start = base_date.replace(year=base_date.year + 1, month=1, day=1)
            end = base_date.replace(year=base_date.year + 1, month=2, day=1) - timedelta(days=1)
        elif base_date.month == 11:
            start = base_date.replace(month=12, day=1)
            end = base_date.replace(year=base_date.year + 1, month=1, day=1) - timedelta(days=1)
        else:
            start = base_date.replace(month=base_date.month + 1, day=1)
            end = base_date.replace(month=base_date.month + 2, day=1) - timedelta(days=1)
        return start, end
    
    elif mode == 'one_year':
        # 미래 1년
        start = base_date
        end = base_date.replace(year=base_date.year + 1)
        return start, end
    
    elif mode == 'past_year':
        # 과거 1년
        start = base_date.replace(year=base_date.year - 1)
        end = base_date
        return start, end
    
    else:
        # 기본값: 이번달
        return get_period_dates('this_month', base_date)


def format_date(date_obj, format_type='display'):
    """
    날짜 객체를 문자열로 포맷
    
    Args:
        date_obj: datetime 객체
        format_type: 'display', 'datetime_display', 'log_format'
        
    Returns:
        포맷된 날짜 문자열
    """
    if date_obj is None or pd.isna(date_obj):
        return ''
    
    formats = {
        'display': '%Y-%m-%d',
        'datetime_display': '%Y-%m-%d %H:%M:%S',
        'log_format': '%Y-%m-%d %H:%M'
    }
    
    fmt = formats.get(format_type, formats['display'])
    
    try:
        return date_obj.strftime(fmt)
    except:
        return str(date_obj)
