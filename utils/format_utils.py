# utils/format_utils.py
"""값 변환 및 포맷 관련 유틸리티 함수"""

import pandas as pd


def safe_int_convert(value, default=0):
    """
    안전한 정수 변환 함수
    
    Args:
        value: 변환할 값
        default: 변환 실패 시 반환할 기본값
        
    Returns:
        int: 변환된 정수 또는 기본값
    """
    if pd.isna(value) or value == '' or value is None:
        return default
    
    try:
        # 문자열인 경우 콤마 제거
        if isinstance(value, str):
            value = value.replace(',', '').strip()
            # 빈 문자열이거나 '-'만 있는 경우 기본값 반환
            if not value or value == '-':
                return default
        
        # int 변환 시도
        return int(float(value))
    except (ValueError, TypeError):
        return default


def safe_float_convert(value, default=0.0):
    """
    안전한 실수 변환 함수
    
    Args:
        value: 변환할 값
        default: 변환 실패 시 반환할 기본값
        
    Returns:
        float: 변환된 실수 또는 기본값
    """
    if pd.isna(value) or value == '' or value is None:
        return default
    
    try:
        # 문자열인 경우 콤마 제거
        if isinstance(value, str):
            value = value.replace(',', '').replace('%', '').strip()
            # 빈 문자열이거나 '-'만 있는 경우 기본값 반환
            if not value or value == '-':
                return default
        
        return float(value)
    except (ValueError, TypeError):
        return default


def format_currency(value, include_currency=False):
    """
    숫자를 천 단위 콤마가 포함된 문자열로 변환
    
    Args:
        value: 숫자 값
        include_currency: '원' 단위 추가 여부
        
    Returns:
        str: 포맷된 문자열
    """
    try:
        num = safe_int_convert(value)
        formatted = f"{num:,}"
        if include_currency:
            formatted += "원"
        return formatted
    except:
        return str(value) if value else "0"


def format_percentage(value, decimals=1):
    """
    숫자를 백분율 문자열로 변환
    
    Args:
        value: 숫자 값 (0.1 = 10%)
        decimals: 소수점 자릿수
        
    Returns:
        str: 포맷된 백분율 문자열
    """
    try:
        num = safe_float_convert(value)
        return f"{num:.{decimals}f}%"
    except:
        return str(value) if value else "0%"


def clean_column_value(value):
    """
    컬럼 값 정리 (공백 제거 등)
    
    Args:
        value: 원본 값
        
    Returns:
        정리된 값
    """
    if value is None:
        return ''
    if isinstance(value, str):
        return value.strip()
    return value
