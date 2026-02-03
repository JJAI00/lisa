# config.py
# LISA 프로젝트 중앙 설정 파일
# 모든 구글 시트 ID 및 공통 설정을 한 곳에서 관리

import os
import sys

def resource_path(relative_path):
    """리소스 파일 경로 반환 (개발/배포 환경 모두 지원)"""
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.abspath('.')
    return os.path.join(base_path, relative_path)


# ========== Google Sheets API 설정 ==========

# 서비스 계정 키 파일 경로
KEY_FILE = resource_path('google_sheet_key/renewal-bot-463605-b8f41de8fbdb.json')

# API 권한 범위
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


# ========== Google Sheet IDs ==========

SHEET_IDS = {
    # 리뉴얼 관리 핵심 데이터
    'renewal_list': '17rfmG5DEOj1CC-iA6SEVIdsS8zP-UC7lGBT6w2LwBvQ',
    
    # 고객사 담당자 정보
    'customer_list': '1hURXlr1g7qMPoj7GxblVC1SDgM9hCwR6qyG20We1188',
    
    # 견적서 및 Status 관리
    'quote_list': '1xjIjFe1Q9dq2zeQOkBSpbqIUOyXAxbyTO8lKN-r1lw0',
    
    # 제품별 가격 정보
    'price_list': '1OGt__Olrempc5ZUy6QqqhvgRseFLanv1xCPaw5Y2qJk',
    
    # 작업 로그
    'log_list': '1llolEKlleT6Cve_M5Qt_CO_JDxQczIKBMqL8fj5vfSk',
    
    # 사용자 인증
    'login_list': '1KEcLicMV6p-RSXl2kg1AHyD2CmU8wABeXDlMTk0gqZY'
}


# ========== 워크시트 이름 ==========

WORKSHEET_NAMES = {
    'renewal_list': 'Sheet1',
    'customer_list': 'Sheet1',
    'quote_list': 'quote_list',  # 또는 'Sheet1' fallback
    'price_list': 'Sheet1',
    'log_list': 'Sheet1',
    'login_list': 'Sheet1'
}


# ========== UI 설정 ==========

# 금액 컬럼 (천 단위 콤마 적용)
COST_COLUMNS = [
    '판매 단가', '판매 합계', '원가', '원가계',
    '실 매입 단가', '실 매입 합계', '이익액', '실 이익액',
    '입금액', '잔액'
]

# 숫자 컬럼 (우측 정렬)
NUMERIC_COLUMNS = COST_COLUMNS + ['이익율', '실 이익율', '수량']


# ========== 캐시 설정 ==========

# 캐시 TTL (초)
CACHE_TTL_SECONDS = 300  # 5분


# ========== 날짜 형식 ==========

DATE_FORMATS = {
    'display': '%Y-%m-%d',
    'datetime_display': '%Y-%m-%d %H:%M:%S',
    'log_format': '%Y-%m-%d %H:%M'
}
