# -*- coding: utf-8 -*-
"""
LISA 프로젝트 상세 매뉴얼 & UI 기획서 생성
각 Python 파일별 UI 화면, 기능, 사용법을 포함
"""

import pandas as pd
from datetime import datetime
import os

# 현재 스크립트의 디렉토리를 작업 디렉토리로 설정
os.chdir(os.path.dirname(os.path.abspath(__file__)))

output_path = 'LISA_상세매뉴얼_UI기획서.xlsx'

print(f"📚 LISA 상세 매뉴얼 & UI 기획서 생성 중: {output_path}")

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    
    # ========== 표지 ==========
    cover_data = {
        '항목': ['문서명', '프로젝트', '버전', '문서 버전', '작성일', '문서 유형', '대상 독자'],
        '내용': [
            'LISA 상세 매뉴얼 & UI 기획서',
            'LISA (License Information System for Administration)',
            'v2.01',
            'v1.0',
            datetime.now().strftime('%Y-%m-%d'),
            '사용자 매뉴얼, UI/UX 기획서',
            '개발자, 사용자, 유지보수 담당자'
        ]
    }
    df_cover = pd.DataFrame(cover_data)
    df_cover.to_excel(writer, sheet_name='📋 표지', index=False)
    
    # ========== 목차 ==========
    toc_data = {
        '번호': list(range(1, 14)),
        '모듈': [
            'renewal_gui.py',
            'renewal_preview.py',
            'customer_info.py',
            'quote_info.py',
            'purchase_sales_info.py',
            'dashboard_info.py',
            'interactive_dashboard.py',
            'team_performance_dashboard.py',
            'license_certificate_info.py',
            'renewal_email.py',
            'renewal_quote.py',
            'lisa_daou_sign.py',
            'data_loader.py'
        ],
        '페이지명': [
            '메인 애플리케이션',
            'Renewal 관리',
            '고객사 관리',
            '견적서 관리',
            '매입/매출 관리',
            '기본 대시보드',
            '인터랙티브 대시보드',
            '팀별 성과 대시보드',
            '인증서 관리',
            '이메일 발송',
            '견적서 생성',
            'DAOU 전자결재 연동',
            '데이터 로더'
        ]
    }
    df_toc = pd.DataFrame(toc_data)
    df_toc.to_excel(writer, sheet_name='📑 목차', index=False)
    
    # ========== 1. renewal_gui.py (메인 애플리케이션) ==========
    main_ui_data = {
        'UI 요소': [
            '🖥️ 메인 윈도우',
            '',
            '',
            '🔐 로그인 창',
            '',
            '',
            '',
            '',
            '📍 사이드바 (좌측)',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '📄 메인 프레임 (우측)',
            '',
            '⏳ 로딩 인디케이터',
            '',
            ''
        ],
        '설명': [
            '전체 애플리케이션 메인 창',
            '크기: 1300x850',
            '배경색: #F7F9FB (연한 회색)',
            '로그인 다이얼로그 (350x250)',
            'ID 입력 필드',
            'PW 입력 필드 (***)',
            '로그인 버튼 (파란색)',
            '구글 시트 기반 인증 (SHA-256)',
            '좌측 네비게이션 메뉴',
            '너비: 200px',
            '배경: #1E293B (어두운 회색)',
            '6개 메뉴 아이템:',
            '  • 🔄 Renewal 관리',
            '  • 👥 고객사 관리',
            '  • 📋 견적서',
            '  • 💰 매입/매출',
            '  • 📊 대시보드 (서브메뉴)',
            '  • 🔐 인증서',
            '탭 컨텐츠 표시 영역',
            '선택된 탭에 따라 동적 변경',
            '데이터 로딩 중 표시',
            'Toplevel 창',
            '프로그레스 바 + 메시지'
        ],
        '기능': [
            'MainApp 클래스',
            '탭 전환 관리',
            '데이터 로더 초기화',
            'LoginWindow 클래스',
            '사용자 ID 입력',
            '비밀번호 입력 (해시 처리)',
            '로그인 버튼 클릭 시 인증',
            '로그인 시트 확인 및 검증',
            '메뉴 클릭 시 탭 전환',
            'Hover 효과 (색상 변경)',
            '선택된 메뉴 하이라이트',
            '',
            'renewal_preview.py',
            'customer_info.py',
            'quote_info.py',
            'purchase_sales_info.py',
            'dashboard_info.py 등',
            'license_certificate_info.py',
            'pack/pack_forget로 전환',
            'Frame 위젯 사용',
            'LoadingIndicator 클래스',
            '백그라운드 데이터 로딩 중 표시',
            '로딩 완료 시 자동 닫기'
        ],
        '사용 방법': [
            'LISA.exe 실행',
            '',
            '',
            '1. 로그인 ID 입력',
            '2. 비밀번호 입력',
            '3. 로그인 버튼 클릭',
            '또는 Enter 키',
            '로그인 실패 시 오류 메시지',
            '좌측 메뉴 클릭하여 탭 이동',
            '마우스 오버 시 색상 변경',
            '클릭 시 해당 탭으로 이동',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '각 탭의 기능 사용',
            '',
            '자동으로 표시됨',
            '로딩 중 기다리기',
            '완료되면 자동 닫힘'
        ]
    }
    df_main = pd.DataFrame(main_ui_data)
    df_main.to_excel(writer, sheet_name='1️⃣ 메인 앱 (renewal_gui)', index=False)
    
    # ========== 2. renewal_preview.py (Renewal 관리) ==========
    renewal_ui_data = {
        'UI 요소': [
            '📌 헤더 섹션',
            '',
            '',
            '🔍 필터 카드 1 (기준 선택)',
            '',
            '',
            '📅 필터 카드 2 (기간 선택)',
            '',
            '',
            '',
            '',
            '',
            '🔎 필터 카드 3 (상세 필터)',
            '',
            '',
            '',
            '',
            '📊 데이터 테이블 (Treeview)',
            '',
            '',
            '',
            '',
            '',
            '🔘 액션 버튼',
            '',
            '',
            '',
            '',
            '',
            '📝 우클릭 메뉴',
            '',
            ''
        ],
        '설명': [
            '페이지 제목: "Renewal 관리"',
            '사용자 정보 표시 (우측)',
            '배경: #F7F9FB',
            '라디오 버튼 2개',
            '  • 만료일 기준',
            '  • 계산서 발행일 기준',
            '기간 선택 버튼',
            '  • 이번달',
            '  • 다음달',
            '  • 이번분기',
            '  • 다음분기',
            '  • 사용자정의 (DateEntry)',
            '필터 드롭다운',
            '  • 고객사명',
            '  • 거래처명',
            '  • 대분류',
            '  • 영업사원',
            '데이터 표시 테이블',
            '컬럼: 고객사명, 제품, 만료일, 판매단가 등',
            '정렬 가능 (헤더 클릭)',
            '다중 선택 가능 (Ctrl/Shift + 클릭)',
            '총 건수, 판매액 통계 표시',
            '색상 코딩 (만료 임박 빨간색)',
            '하단 액션 버튼 영역',
            '  • 이메일 작성',
            '  • 견적서+메일',
            '  • 여러 제품 선택 이메일',
            '  • 여러 제품 선택 견적서',
            '  • 로그 조회',
            '테이블 행 우클릭 시',
            '  • 고객사 상세 정보',
            '  • DAOU 연동'
        ],
        '주요 기능': [
            '제목 표시, 사용자 이름 표시',
            '',
            '',
            '만료일 또는 계산서발행일 선택',
            '선택된 기준으로 데이터 필터링',
            '',
            '버튼 클릭으로 빠른 기간 선택',
            'load_this_month()',
            'load_next_month()',
            'load_this_quarter()',
            'load_next_quarter()',
            'load_custom_period()',
            '드롭다운으로 상세 필터링',
            '고객사명 선택',
            '거래처명 선택',
            '대분류 선택',
            '영업사원 선택',
            'Treeview 위젯',
            'renewal_list 데이터 표시',
            '헤더 클릭으로 정렬',
            'Ctrl/Shift로 다중 선택',
            '통계 자동 계산',
            '만료일 임박 하이라이트',
            '선택된 행에 대해 액션 실행',
            'send_emails()',
            'create_and_send_quote()',
            'send_multi_emails()',
            'create_multi_quotes()',
            'show_logs()',
            '추가 기능 메뉴',
            'show_customer_detail()',
            'open_daou_automation()'
        ],
        '사용 시나리오': [
            '탭 선택 시 자동 표시',
            '',
            '',
            '1. 기준 선택 (만료일 권장)',
            '',
            '',
            '2. 기간 선택',
            '  - 이번달: 현재 월 데이터',
            '  - 다음달: 다음 월 데이터',
            '  - 이번분기: 현재 분기',
            '  - 다음분기: 다음 분기',
            '  - 사용자정의: 날짜 직접 선택',
            '3. 필요시 상세 필터 적용',
            '  - 특정 고객사만 보기',
            '  - 특정 제품군만 보기',
            '  - 담당자별 보기',
            '',
            '4. 데이터 확인',
            '  - 만료 임박 항목 확인',
            '  - 헤더 클릭으로 정렬',
            '  - 행 선택 (단일/다중)',
            '',
            '',
            '5. 액션 실행',
            '  - 선택 후 버튼 클릭',
            '  - Outlook 자동 실행',
            '  - 템플릿 기반 이메일 작성',
            '  - 견적서 자동 생성',
            '  - 로그 자동 기록',
            '6. 추가 기능',
            '  - 우클릭으로 고객사 상세',
            '  - DAOU 연동으로 전자결재'
        ]
    }
    df_renewal = pd.DataFrame(renewal_ui_data)
    df_renewal.to_excel(writer, sheet_name='2️⃣ Renewal 관리', index=False)
    
    # ========== 3. customer_info.py (고객사 관리) ==========
    customer_ui_data = {
        'UI 요소': [
            '📌 헤더',
            '',
            '🔍 필터 섹션',
            '',
            '',
            '',
            '📊 고객사 테이블',
            '',
            '',
            '',
            '📈 통계 카드',
            '',
            '',
            '🔘 액션 버튼',
            '',
            ''
        ],
        '설명': [
            '제목: "고객사 관리"',
            '사용자 정보',
            '기간 선택 (이번달/다음달/분기)',
            '고객사명 필터',
            '거래처명 필터',
            '대분류/영업사원 필터',
            '고객사별 데이터 표시',
            '컬럼: 고객사명, 제품, 판매액, 이익액 등',
            '고객사 중복 제거',
            '고객사별 거래 내역 집계',
            '총 거래 건수',
            '총 판매액',
            '총 이익액',
            '신규 고객사 등록',
            '고객사 정보 수정',
            '고객사 상세 정보'
        ],
        '주요 기능': [
            'CustomerInfoGUI 클래스',
            '',
            'load_period() 메서드',
            '고객사명 드롭다운',
            '거래처명 드롭다운',
            '대분류/영업사원 드롭다운',
            'Treeview 위젯',
            'renewal_list + customer_list 통합',
            '고객사별 그룹화',
            '거래 내역 집계',
            '자동 계산',
            '천 단위 콤마 표시',
            '통계 자동 업데이트',
            'CustomerAddWindow()',
            'update_customer_info()',
            'show_customer_detail()'
        ],
        '사용 방법': [
            '고객사 관리 탭 클릭',
            '',
            '1. 기간 선택',
            '2. 필요시 필터 적용',
            '3. 특정 고객사 검색',
            '4. 분류별 필터',
            '5. 고객사 목록 확인',
            '   - 고객사별 거래 건수',
            '   - 판매액, 이익액 확인',
            '   - 정렬하여 분석',
            '6. 통계 확인',
            '   - 전체 거래 건수',
            '   - 총 판매액',
            '7. 신규 고객사 등록',
            '   - 버튼 클릭',
            '   - 정보 입력 후 저장'
        ]
    }
    df_customer = pd.DataFrame(customer_ui_data)
    df_customer.to_excel(writer, sheet_name='3️⃣ 고객사 관리', index=False)
    
    # ========== 4. quote_info.py (견적서 관리) ==========
    quote_ui_data = {
        'UI 요소': [
            '📌 헤더',
            '🔍 필터 섹션',
            '',
            '',
            '📊 견적서 테이블',
            '',
            '',
            '',
            '',
            '🎯 Status 관리',
            '',
            '',
            '',
            '',
            '🔘 액션 버튼',
            '',
            '',
            ''
        ],
        '설명': [
            '제목: "견적서 관리"',
            '일자 기준 기간 선택',
            '고객사명/거래처명 필터',
            'Status 필터',
            '견적서 데이터 표시',
            'ID, FCST, Status, 일자',
            '고객사명, 제품, 판매 합계',
            '원가계, 이익액, 이익율',
            'Status별 색상 코딩',
            'Status 컬럼 (10%~100%)',
            '진행률 표시',
            '색상으로 구분:',
            '  10%~70%: 회색',
            '  80%~90%: 노란색',
            '  100%: 초록색',
            '신규 견적서',
            '견적서 수정',
            'Status 일괄 수정',
            'DAOU 연동'
        ],
        '주요 기능': [
            'QuoteInfoGUI 클래스',
            'load_period() 메서드',
            '드롭다운 필터',
            'Status 필터링',
            'quote_list 데이터 표시',
            '자동 ID 생성',
            'FCST (예상 발주일)',
            'Status 관리',
            'Status별 색상',
            'QuoteStatusModifier',
            '10%, 20%, ..., 100%',
            'status_to_color()',
            '',
            '',
            '',
            'QuoteNewWindow()',
            'QuoteModifyWindow()',
            'show_status_context_menu()',
            'lisa_daou_sign()'
        ],
        '사용 방법': [
            '견적서 탭 클릭',
            '1. 기간 선택 (일자 기준)',
            '2. 필요시 필터 적용',
            '3. Status로 필터',
            '4. 견적서 목록 확인',
            '   - Status 진행률 확인',
            '   - 색상으로 빠른 파악',
            '   - 판매액, 이익액 확인',
            '',
            '5. Status 관리',
            '   - 우클릭하여 Status 변경',
            '   - 10% 단위로 진행률 관리',
            '   - 100% = 계약 완료',
            '',
            '',
            '6. 신규 견적서 작성',
            '7. 기존 견적서 수정',
            '8. Status 일괄 변경',
            '9. DAOU 전자결재 연동'
        ]
    }
    df_quote = pd.DataFrame(quote_ui_data)
    df_quote.to_excel(writer, sheet_name='4️⃣ 견적서 관리', index=False)
    
    # ========== 5. purchase_sales_info.py (매입/매출 관리) ==========
    purchase_ui_data = {
        'UI 요소': [
            '📌 헤더',
            '🔍 필터 섹션',
            '',
            '📊 매입/매출 테이블',
            '',
            '',
            '',
            '📈 통계 섹션',
            '',
            '',
            '',
            '🔘 액션 버튼',
            '',
            ''
        ],
        '설명': [
            '제목: "매입/매출 관리"',
            '계산서 발행일 기준 기간 선택',
            '고객사/거래처/대분류/영업사원 필터',
            'renewal_list 전체 컬럼 표시',
            '판매 단가, 원가, 이익액, 이익율',
            '실 매입 단가, 실 매입 합계',
            '실 이익액, 실 이익율',
            '총 판매액 합계',
            '총 원가 합계',
            '총 이익액 합계',
            '평균 이익율',
            '매입/매출 정보 수정',
            '만료일 일괄 수정',
            'Excel 내보내기'
        ],
        '주요 기능': [
            'PurchaseSalesInfo 클래스',
            'load_period() 메서드',
            '다중 필터 조합',
            'renewal_list 데이터',
            '수익성 분석',
            '실제 매입가 vs 예상 매입가',
            '실제 이익 vs 예상 이익',
            '자동 계산',
            '천 단위 콤마',
            'sum() 함수',
            'mean() 함수',
            'PurchaseSalesModifyDialog()',
            'PurchaseExpDateModify()',
            'export_to_excel()'
        ],
        '사용 방법': [
            '매입/매출 탭 클릭',
            '1. 기간 선택 (계산서발행일)',
            '2. 필터로 조건 설정',
            '3. 매입/매출 데이터 확인',
            '   - 판매 단가 확인',
            '   - 원가 확인',
            '   - 이익액, 이익율 분석',
            '4. 통계 확인',
            '   - 총 판매액',
            '   - 총 원가',
            '   - 총 이익액',
            '5. 데이터 수정',
            '   - 행 선택 후 수정',
            '   - 만료일 일괄 변경'
        ]
    }
    df_purchase = pd.DataFrame(purchase_ui_data)
    df_purchase.to_excel(writer, sheet_name='5️⃣ 매입매출 관리', index=False)
    
    # ========== 6. 대시보드들 ==========
    dashboard_ui_data = {
        '대시보드 종류': [
            '📊 기본 대시보드',
            '(dashboard_info.py)',
            '',
            '',
            '',
            '',
            '📊 인터랙티브 대시보드',
            '(interactive_dashboard.py)',
            '',
            '',
            '',
            '',
            '',
            '📊 팀별 성과 대시보드',
            '(team_performance_dashboard.py)',
            '',
            '',
            '',
            ''
        ],
        'UI 구성': [
            '월 선택 드롭다운',
            '영업사원 선택',
            'Status 분포 파이차트',
            '영업사원별 매출 막대차트',
            'matplotlib 차트',
            '',
            '조직 필터 (부서/팀/영업사원)',
            '차트 타입 선택 (라디오버튼)',
            '  - Status 분포',
            '  - 월별 추이',
            '  - 부서별 성과',
            '  - 팀별 성과',
            '  - 영업사원별 성과',
            '  - 제품별 분석',
            '연도/월/분기 선택',
            '부서/팀/영업사원 필터',
            'Target vs 실적 비교',
            '조직별 성과 분석',
            '캐싱 (30분 TTL)'
        ],
        '주요 기능': [
            'DashboardDialog 클래스',
            'quote_list 데이터 사용',
            'Status별 집계',
            '영업사원별 집계',
            '한글 폰트 설정',
            '',
            'InteractiveDashboard 클래스',
            'OrganizationManager 연동',
            '파이차트',
            '선그래프 (월별 추이)',
            '막대차트 (부서/팀/영업사원)',
            '막대차트 (제품별)',
            '드릴다운 기능',
            '차트 히스토리',
            'TeamPerformanceDashboard 클래스',
            'Target 워크시트 연동',
            'Target 대비 달성률',
            '조직 계층 구조 분석',
            'pickle 캐싱'
        ],
        '사용 시나리오': [
            '1. 대시보드 탭 선택',
            '2. 월 선택',
            '3. Status 분포 확인',
            '4. 영업사원별 실적 비교',
            '5. 차트 분석',
            '',
            '1. 대시보드 > 인터랙티브',
            '2. 조직 필터 설정',
            '3. 차트 타입 선택',
            '4. 차트 확인',
            '5. 클릭하여 드릴다운',
            '6. 상세 데이터 확인',
            '7. 뒤로가기로 복귀',
            '',
            '1. 대시보드 > 팀별',
            '2. 연도/월 선택',
            '3. Target 대비 실적 확인',
            '4. 조직별 성과 분석',
            '5. 캐시된 데이터 활용'
        ]
    }
    df_dashboard = pd.DataFrame(dashboard_ui_data)
    df_dashboard.to_excel(writer, sheet_name='6️⃣ 대시보드 (3종)', index=False)
    
    # ========== 7. license_certificate_info.py (인증서 관리) ==========
    cert_ui_data = {
        'UI 요소': [
            '📌 헤더',
            '🎨 인증서 버튼 (5개)',
            '',
            '',
            '',
            '',
            '각 버튼별 다이얼로그',
            '',
            '',
            ''
        ],
        '설명': [
            '제목: "인증서 관리"',
            '5개 제조사 버튼 (2열 그리드)',
            '  • BorisFX (보라색)',
            '  • Foundry (파란색)',
            '  • Marmoset (초록색)',
            '  • Maxon (주황색)',
            '  • VideoCopilot (빨간색)',
            '버튼 클릭 시 다이얼로그',
            '고객사 정보 입력',
            '인증서 자동 생성'
        ],
        '주요 기능': [
            'open_dialog() 함수',
            '5개 제조사별 모듈 호출',
            'open_borisfx_dialog()',
            'open_foundry_dialog()',
            'open_marmoset_dialog()',
            'open_maxon_dialog()',
            'open_videocopilot_dialog()',
            '각 다이얼로그:',
            '  - 고객사명 입력',
            '  - 제품 선택',
            '  - 수량 입력',
            '  - 만료일 선택',
            '  - 템플릿 자동 채우기',
            '  - PDF 변환 (선택사항)',
            '  - 저장 위치 선택'
        ],
        '사용 방법': [
            '인증서 탭 클릭',
            '1. 제조사 버튼 선택',
            '2. 다이얼로그에서 정보 입력',
            '   - 고객사명',
            '   - 제품명',
            '   - 수량',
            '   - 만료일',
            '3. 생성 버튼 클릭',
            '4. 자동으로 템플릿 채우기',
            '   - Word/Excel 자동화',
            '   - 셀/필드 자동 입력',
            '5. PDF 변환 (선택)',
            '6. 파일 저장',
            '   - license_certificate/ 폴더',
            '   - 고객사명_제조사_날짜.확장자',
            '7. 완료 메시지'
        ]
    }
    df_cert = pd.DataFrame(cert_ui_data)
    df_cert.to_excel(writer, sheet_name='7️⃣ 인증서 관리', index=False)
    
    # ========== 8. renewal_email.py (이메일 발송) ==========
    email_func_data = {
        '함수/기능': [
            '📧 send_emails()',
            '',
            '',
            '',
            '',
            '📧 send_multi_emails()',
            '',
            '',
            '',
            '📝 HTML 템플릿',
            '',
            '',
            '',
            '',
            '📋 로그 기록',
            '',
            ''
        ],
        '설명': [
            '단일 제품 이메일 발송',
            'Outlook COM 객체 사용',
            'HTML 템플릿 로드',
            '템플릿 변수 치환',
            'Display() - 이메일 창 표시',
            '다중 제품 이메일 발송',
            '여러 제품을 하나의 이메일로',
            '제품 목록 테이블 생성',
            'HTML 동적 생성',
            '제품별 템플릿',
            '  - email_template.html (기본)',
            '  - email_template_Chaosgroup.html',
            '  - email_template_Foundry.html',
            '  - email_template_TeamViewer.html',
            '구글 시트 log_list에 기록',
            '일시, 고객사명, 제품, 액션',
            'append_log() 함수'
        ],
        '템플릿 변수': [
            '{{고객사명}}',
            '{{제품}}',
            '{{수량}}',
            '{{만료일}}',
            '{{담당자명}}',
            '{{연락처}}',
            '{{이메일}}',
            '{{주소}}',
            '{{사업자번호}}',
            'HTML 파일에서 변수 사용',
            'Python에서 replace()로 치환',
            '',
            '',
            '',
            '자동 로그 기록',
            '실행자: socket.gethostname()',
            '사용자: 로그인 ID'
        ],
        '사용 흐름': [
            '1. Renewal 관리 탭에서 행 선택',
            '2. "이메일 작성" 버튼 클릭',
            '3. Outlook 확인',
            '   - Outlook이 실행 중이어야 함',
            '   - COM 객체 생성',
            '4. 템플릿 로드',
            '   - 대분류별 템플릿 선택',
            '   - 없으면 기본 템플릿',
            '',
            '5. 템플릿 변수 치환',
            '   - 고객사 정보 입력',
            '   - 제품 정보 입력',
            '   - 만료일 등 입력',
            '',
            '6. Outlook 창 표시',
            '   - 검토 후 전송',
            '   - 로그 자동 기록'
        ]
    }
    df_email = pd.DataFrame(email_func_data)
    df_email.to_excel(writer, sheet_name='8️⃣ 이메일 발송', index=False)
    
    # ========== 9. renewal_quote.py (견적서 생성) ==========
    quote_func_data = {
        '함수/기능': [
            '📋 create_and_send_quote()',
            '',
            '',
            '',
            '',
            '',
            '📋 create_multi_quotes()',
            '',
            '',
            '📄 Excel 템플릿',
            '',
            '',
            '',
            '📄 PDF 변환',
            '',
            ''
        ],
        '설명': [
            '단일 견적서 생성',
            'Excel COM 객체 사용',
            '템플릿 파일 열기',
            '셀 데이터 입력',
            'PDF로 변환',
            'Outlook에 첨부',
            '다중 견적서 생성',
            '여러 제품을 하나의 견적서로',
            '제품 행 동적 추가',
            'quote/quote_origin_01.xlsx',
            'Excel 템플릿 파일',
            '셀 위치 정의:',
            '  - E3: 일자',
            '  - C4: 고객사명',
            '  - E6: 담당자명',
            '  - A10~: 제품 목록',
            'ExportAsFixedFormat(0, path)',
            'PDF 형식으로 저장',
            '이메일 첨부용'
        ],
        '템플릿 구조': [
            'A열: 번호',
            'B열: 품명',
            'C열: 규격',
            'D열: 수량',
            'E열: 단가',
            'F열: 공급가액',
            '',
            '여러 제품은 행 추가',
            '합계 자동 계산',
            'Excel 수식 사용',
            '=SUM(F10:F20)',
            '',
            '',
            '',
            '',
            ''
        ],
        '사용 흐름': [
            '1. Renewal 관리에서 행 선택',
            '2. "견적서+메일" 버튼 클릭',
            '3. Excel 자동 실행',
            '   - 템플릿 열기',
            '   - 데이터 입력',
            '   - 제품 정보 채우기',
            '4. PDF 변환',
            '   - 임시 폴더에 저장',
            '   - quote/ 폴더',
            '5. Outlook 실행',
            '   - 견적서 PDF 첨부',
            '   - 이메일 본문 작성',
            '6. Display()',
            '   - 이메일 창 표시',
            '   - 검토 후 전송',
            '   - 로그 기록'
        ]
    }
    df_quote_func = pd.DataFrame(quote_func_data)
    df_quote_func.to_excel(writer, sheet_name='9️⃣ 견적서 생성', index=False)
    
    # ========== 10. lisa_daou_sign.py (DAOU 연동) ==========
    daou_func_data = {
        '클래스/기능': [
            '🤖 LisaDaouSign 클래스',
            '',
            '',
            '',
            '🌐 setup_driver()',
            '',
            '',
            '🔐 login()',
            '',
            '',
            '📝 create_document()',
            '',
            '',
            '📎 upload_file()',
            '',
            '✅ submit()',
            '',
            '⚙️ 선택자 관리',
            '',
            ''
        ],
        '설명': [
            'Selenium 기반 자동화',
            'ChromeDriver 사용',
            '헤드리스 모드 (보이지 않음)',
            '또는 일반 모드 (디버깅용)',
            '크롬 드라이버 설정',
            'webdriver-manager로 자동 설치',
            '기존 프로필 사용 (로그인 유지)',
            'DAOU 오피스 로그인',
            '저장된 크롬 프로필 사용',
            '또는 ID/PW 자동 입력',
            '결재 문서 작성',
            '제목, 내용 입력',
            '결재선 선택',
            '첨부 파일 업로드',
            '견적서 PDF 등',
            '결재 요청 제출',
            '자동으로 클릭',
            'lisa_daou_selectors.json',
            'UI 선택자 저장',
            '선택자 변경 시 자동 업데이트'
        ],
        '자동화 단계': [
            '1. Selenium 초기화',
            '2. ChromeDriver 설정',
            '3. 헤드리스 옵션',
            '4. 프로필 로드',
            '5. 브라우저 실행',
            '',
            '',
            '6. DAOU 사이트 접속',
            '7. 로그인 확인',
            '8. 세션 유지',
            '9. 결재 페이지 이동',
            '10. 문서 양식 선택',
            '11. 제목/내용 입력',
            '12. 파일 업로드',
            '13. 업로드 대기',
            '14. 결재선 선택',
            '15. 제출 버튼 클릭',
            '16. 완료 확인',
            '17. 로그 기록',
            '18. 브라우저 종료'
        ],
        '사용 방법': [
            'Renewal 관리 또는 견적서 관리에서',
            '우클릭 메뉴 > "DAOU 연동"',
            '',
            '',
            '1. DAOU 로그인 확인',
            '   - 크롬에 로그인되어 있어야 함',
            '',
            '2. 자동화 시작',
            '   - 백그라운드 실행',
            '   - 진행 상황 로그 출력',
            '3. 문서 작성',
            '   - 자동으로 정보 입력',
            '   - 견적서 첨부',
            '4. 결재 요청',
            '   - 자동 제출',
            '5. 완료',
            '   - 성공/실패 메시지',
            '6. 오류 발생 시',
            '   - 로그 확인',
            '   - 선택자 업데이트 필요'
        ]
    }
    df_daou = pd.DataFrame(daou_func_data)
    df_daou.to_excel(writer, sheet_name='🔟 DAOU 전자결재', index=False)
    
    # ========== 11. data_loader.py (데이터 로더) ==========
    loader_data = {
        '클래스/기능': [
            '💾 DataLoader 클래스',
            '',
            '',
            '🔄 load_current_month_data()',
            '',
            '🔄 load_all_data()',
            '',
            '📊 get_renewal_data()',
            '📊 get_customer_data()',
            '📊 get_quote_data()',
            '📊 get_price_data()',
            '📊 get_log_data()',
            '💿 캐싱 시스템',
            '',
            '',
            '🔁 safe_api_call()',
            '',
            ''
        ],
        '설명': [
            '구글 시트 데이터 관리',
            'gspread 사용',
            '5개 워크시트 연동',
            '초기 로딩 (빠른 시작)',
            '현재 월 데이터만',
            '전체 데이터 로딩',
            '백그라운드 스레드',
            'renewal_list 반환',
            'customer_list 반환',
            'quote_list 반환',
            'price_list 반환',
            'log_list 반환',
            '5분 TTL',
            '메모리 캐싱',
            'pickle 파일 캐싱 (선택)',
            'API 호출 래퍼',
            '재시도 로직 (지수 백오프)',
            '429/503 에러 핸들링'
        ],
        '캐시 전략': [
            '1. 메모리 캐시 확인',
            '2. 캐시 있으면 반환',
            '3. 캐시 없으면 API 호출',
            '4. 초기 로딩',
            '   - 현재 월만',
            '5. 백그라운드 로딩',
            '   - 전체 데이터',
            '6. 캐시 키 생성',
            '   - renewal_data',
            '   - customer_data',
            '   - quote_data',
            '   - price_data',
            '   - log_data',
            '7. TTL 5분',
            '   - 5분 경과 시 재로드',
            '   - API 호출 최소화',
            '8. 재시도 로직',
            '   - 최대 3회',
            '   - 지수 백오프 (1s, 2s, 4s)'
        ],
        '성능 최적화': [
            '초기 로딩 속도 향상',
            '현재 월 데이터만 먼저 로드',
            '사용자는 바로 사용 가능',
            '',
            '',
            '백그라운드 전체 로딩',
            '사용 중에 전체 데이터 로드',
            'DataFrame 재사용',
            'get 메서드로 캐시 반환',
            '',
            '',
            '',
            '메모리 캐싱',
            '빠른 접근',
            'API 호출 감소',
            'API 호출 최적화',
            '재시도 로직',
            '에러 복구'
        ]
    }
    df_loader = pd.DataFrame(loader_data)
    df_loader.to_excel(writer, sheet_name='1️⃣1️⃣ 데이터 로더', index=False)
    
    # ========== 사용자 시나리오 ==========
    scenario_data = {
        '시나리오': [
            '📧 리뉴얼 이메일 발송',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '📋 견적서 생성 및 발송',
            '',
            '',
            '',
            '',
            '',
            '',
            '👥 고객사 정보 조회',
            '',
            '',
            '',
            '',
            '💰 수익성 분석',
            '',
            '',
            '',
            '',
            '📊 월별 실적 확인',
            '',
            '',
            '',
            '🔐 인증서 발급',
            '',
            '',
            ''
        ],
        '단계별 절차': [
            '1. LISA 실행 및 로그인',
            '2. Renewal 관리 탭 선택',
            '3. 만료일 기준 선택',
            '4. 이번달 버튼 클릭',
            '5. 만료 임박 고객 확인 (빨간색)',
            '6. 이메일 보낼 고객 선택',
            '7. "이메일 작성" 버튼 클릭',
            '8. Outlook에서 검토 후 전송',
            '1. Renewal 관리 탭',
            '2. 견적서 필요 고객 선택',
            '3. "견적서+메일" 버튼 클릭',
            '4. Excel 견적서 자동 생성',
            '5. PDF 변환',
            '6. Outlook에 자동 첨부',
            '7. 이메일 검토 후 전송',
            '1. 고객사 관리 탭 선택',
            '2. 기간 선택 (이번달)',
            '3. 특정 고객사명으로 필터',
            '4. 거래 내역 확인',
            '5. 통계 확인 (판매액, 이익액)',
            '1. 매입/매출 관리 탭',
            '2. 기간 선택 (월/분기)',
            '3. 영업사원별 필터',
            '4. 판매 단가, 원가, 이익율 확인',
            '5. 통계로 수익성 분석',
            '1. 대시보드 탭 선택',
            '2. 인터랙티브 대시보드',
            '3. 월별 추이 차트 선택',
            '4. 부서/팀별 필터 적용',
            '5. 차트로 실적 확인',
            '1. 인증서 관리 탭',
            '2. 제조사 버튼 선택 (예: Foundry)',
            '3. 고객사 정보 입력',
            '4. 생성 버튼 클릭',
            '5. 자동 생성 완료'
        ],
        '예상 소요 시간': [
            '전체: 약 5분',
            '1분',
            '10초',
            '10초',
            '1분',
            '30초',
            '10초',
            '2분',
            '전체: 약 7분',
            '1분',
            '30초',
            '10초',
            '1분',
            '1분',
            '3분',
            '전체: 약 3분',
            '10초',
            '10초',
            '30초',
            '1분',
            '1분',
            '전체: 약 4분',
            '10초',
            '10초',
            '30초',
            '2분',
            '1분',
            '전체: 약 2분',
            '10초',
            '10초',
            '30초',
            '1분',
            '',
            '전체: 약 3분',
            '10초',
            '10초',
            '1분',
            '30초',
            '1분'
        ]
    }
    df_scenario = pd.DataFrame(scenario_data)
    df_scenario.to_excel(writer, sheet_name='🎬 사용자 시나리오', index=False)
    
    # ========== 문제 해결 가이드 ==========
    troubleshooting_data = {
        '문제': [
            '❌ 로그인 실패',
            '',
            '',
            '❌ 데이터가 안 보임',
            '',
            '',
            '❌ Outlook 오류',
            '',
            '',
            '❌ Excel 오류',
            '',
            '',
            '❌ DAOU 연동 실패',
            '',
            '',
            '❌ 인증서 생성 실패',
            '',
            '',
            '❌ 느린 로딩',
            '',
            ''
        ],
        '원인': [
            'ID/PW 불일치',
            '구글 시트 접근 권한 없음',
            '네트워크 문제',
            '캐시 문제',
            '구글 시트 API 오류',
            '권한 문제',
            'Outlook이 실행되지 않음',
            'COM 객체 오류',
            '템플릿 파일 없음',
            'Excel이 실행되지 않음',
            'COM 객체 오류',
            '템플릿 파일 없음',
            '로그인 세션 만료',
            '선택자 변경',
            '네트워크 문제',
            'Word/Excel 없음',
            '템플릿 파일 없음',
            '권한 문제',
            '전체 데이터 로딩 중',
            'API 호출 제한',
            '네트워크 느림'
        ],
        '해결 방법': [
            '1. ID/PW 확인',
            '2. 구글 시트 공유 확인',
            '3. 인터넷 연결 확인',
            '1. 새로고침 (F5)',
            '2. LISA 재시작',
            '3. 구글 시트 권한 확인',
            '1. Outlook 먼저 실행',
            '2. Outlook 재시작',
            '3. 템플릿 파일 확인',
            '1. Excel 먼저 실행',
            '2. Excel 재시작',
            '3. 템플릿 파일 확인 (quote/)',
            '1. Chrome에서 DAOU 로그인',
            '2. 선택자 JSON 확인',
            '3. 로그 파일 확인',
            '1. Word/Excel 설치 확인',
            '2. 템플릿 파일 확인',
            '3. license_certificate/ 폴더 확인',
            '1. 초기 로딩 대기 (현재 월)',
            '2. 백그라운드 로딩 완료 대기',
            '3. 캐시 활용 (2번째부터 빠름)'
        ]
    }
    df_troubleshooting = pd.DataFrame(troubleshooting_data)
    df_troubleshooting.to_excel(writer, sheet_name='🔧 문제 해결', index=False)
    
    # 열 너비 자동 조정
    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
        for column in worksheet.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 120)
            worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

print(f"✅ LISA 상세 매뉴얼 & UI 기획서 생성 완료!")
print()
print("📚 생성된 시트:")
print("   1. 📋 표지")
print("   2. 📑 목차")
print("   3. 1️⃣ 메인 앱 (renewal_gui)")
print("   4. 2️⃣ Renewal 관리")
print("   5. 3️⃣ 고객사 관리")
print("   6. 4️⃣ 견적서 관리")
print("   7. 5️⃣ 매입매출 관리")
print("   8. 6️⃣ 대시보드 (3종)")
print("   9. 7️⃣ 인증서 관리")
print("  10. 8️⃣ 이메일 발송")
print("  11. 9️⃣ 견적서 생성")
print("  12. 🔟 DAOU 전자결재")
print("  13. 1️⃣1️⃣ 데이터 로더")
print("  14. 🎬 사용자 시나리오")
print("  15. 🔧 문제 해결")
print()
print(f"📂 저장 위치: {os.path.abspath(output_path)}")

