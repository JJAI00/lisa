"""
LISA_재현_프롬프트.md를 기획서 형태의 Excel 문서로 변환
"""

import pandas as pd
from datetime import datetime
import re

def parse_md_file(md_path):
    """마크다운 파일을 파싱하여 구조화된 데이터로 변환"""
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    sections = {}
    current_section = None
    current_subsection = None
    current_content = []
    
    lines = content.split('\n')
    
    for line in lines:
        # H2 헤딩 (## )
        if line.startswith('## '):
            if current_section:
                if current_subsection:
                    if current_section not in sections:
                        sections[current_section] = {}
                    sections[current_section][current_subsection] = '\n'.join(current_content)
                else:
                    sections[current_section] = '\n'.join(current_content)
            
            current_section = line.replace('## ', '').strip()
            current_subsection = None
            current_content = []
        
        # H3 헤딩 (### )
        elif line.startswith('### '):
            if current_subsection:
                if current_section not in sections:
                    sections[current_section] = {}
                sections[current_section][current_subsection] = '\n'.join(current_content)
            
            current_subsection = line.replace('### ', '').strip()
            current_content = []
        
        else:
            current_content.append(line)
    
    # 마지막 섹션 저장
    if current_section:
        if current_subsection:
            if current_section not in sections:
                sections[current_section] = {}
            sections[current_section][current_subsection] = '\n'.join(current_content)
        else:
            sections[current_section] = '\n'.join(current_content)
    
    return sections

def create_excel_planning_doc(md_path, output_path):
    """마크다운을 Excel 기획서로 변환"""
    
    print(f"📖 마크다운 파일 읽기: {md_path}")
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        
        # 1. 표지 시트
        create_cover_sheet(writer)
        
        # 2. 목차 시트
        create_toc_sheet(writer)
        
        # 3. 프로젝트 개요 시트
        create_overview_sheet(writer)
        
        # 4. 시스템 아키텍처 시트
        create_architecture_sheet(writer)
        
        # 5. 핵심 모듈 시트 (6개 탭)
        create_modules_sheet(writer)
        
        # 6. 데이터베이스 구조 시트
        create_database_sheet(writer)
        
        # 7. 외부 연동 시트
        create_integration_sheet(writer)
        
        # 8. UI/UX 디자인 시트
        create_ui_design_sheet(writer)
        
        # 9. 개발 가이드 시트
        create_development_guide_sheet(writer)
        
        # 10. 기술 스택 시트
        create_tech_stack_sheet(writer)
        
        # 11. 배포 가이드 시트
        create_deployment_sheet(writer)
        
        # 12. 체크리스트 시트
        create_checklist_sheet(writer)
    
    print(f"✅ Excel 기획서 생성 완료: {output_path}")

def create_cover_sheet(writer):
    """표지 시트"""
    cover_data = {
        '': ['', '', '', '', '', '', '', '', '', ''],
        ' ': ['', '', '', '📋 LISA 프로젝트 기획서', '', '', 
              '(License Information System for Administration)', '', '', ''],
        '  ': ['', '', '', '', '', '', '', '', '', '']
    }
    
    df = pd.DataFrame(cover_data)
    df.to_excel(writer, sheet_name='📋 표지', index=False, header=False)
    
    worksheet = writer.sheets['📋 표지']
    
    # 제목 스타일 (중앙 정렬, 큰 글씨)
    worksheet.merge_cells('B4:E4')
    worksheet['B4'].font = worksheet['B4'].font.copy(size=24, bold=True)
    worksheet['B4'].alignment = worksheet['B4'].alignment.copy(horizontal='center', vertical='center')
    
    worksheet.merge_cells('B7:E7')
    worksheet['B7'].font = worksheet['B7'].font.copy(size=14)
    worksheet['B7'].alignment = worksheet['B7'].alignment.copy(horizontal='center')
    
    # 하단 정보
    info_data = {
        '항목': ['프로젝트명', '버전', '문서 버전', '작성일', '작성자', '상태'],
        '내용': ['LISA', 'v2.01', 'v1.0', datetime.now().strftime('%Y-%m-%d'), 
                'QVREX 소프트웨어사업1팀', '운영 중']
    }
    
    df_info = pd.DataFrame(info_data)
    # 15행부터 시작
    df_info.to_excel(writer, sheet_name='📋 표지', index=False, startrow=14, startcol=1)

def create_toc_sheet(writer):
    """목차 시트"""
    toc_data = {
        '번호': ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12'],
        '장': ['프로젝트 개요', '시스템 아키텍처', '핵심 모듈', '데이터베이스 구조', 
               '외부 시스템 연동', 'UI/UX 디자인', '개발 가이드', '기술 스택',
               '배포 가이드', '문제 해결', '체크리스트', '부록'],
        '페이지': ['3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14'],
        '설명': [
            '프로젝트 목적, 범위, 주요 기능',
            '메인 애플리케이션, 데이터 관리, 모듈 구조',
            '6개 탭 상세 설명 (Renewal, 고객사, 견적서, 매입/매출, 대시보드, 인증서)',
            '구글 시트 5개 워크시트 구조',
            '구글 시트 API, DAOU 전자결재, Outlook 이메일',
            'Material Design, 색상 팔레트, 컴포넌트',
            'Step by Step 개발 가이드, 코드 예시',
            'Python, Tkinter, pandas, gspread 등',
            'PyInstaller 패키징, 배포 절차',
            '일반적인 오류 및 해결 방법',
            '개발 완료 항목, 테스트 항목',
            '참고 자료, 최종 프롬프트'
        ]
    }
    
    df = pd.DataFrame(toc_data)
    df.to_excel(writer, sheet_name='📑 목차', index=False)
    
    worksheet = writer.sheets['📑 목차']
    worksheet.column_dimensions['A'].width = 8
    worksheet.column_dimensions['B'].width = 25
    worksheet.column_dimensions['C'].width = 10
    worksheet.column_dimensions['D'].width = 70

def create_overview_sheet(writer):
    """프로젝트 개요 시트"""
    overview_data = {
        '구분': [
            '프로젝트명',
            '영문명',
            '버전',
            '개발 기간',
            '개발 언어',
            'UI 프레임워크',
            '개발 조직',
            '',
            '프로젝트 목적',
            '',
            '',
            '주요 기능',
            '',
            '',
            '',
            '',
            '',
            '',
            '기대 효과',
            '',
            '',
            ''
        ],
        '내용': [
            'LISA',
            'License Information System for Administration',
            'v2.01',
            '2024-2025',
            'Python 3.x',
            'Tkinter (Material Design 스타일)',
            '큐브렉스(QVREX) 소프트웨어사업1팀',
            '',
            '소프트웨어 라이선스 리뉴얼 관리 및 영업 지원 자동화',
            '고객사 정보 통합 관리',
            '매입/매출 수익성 분석',
            '',
            '1. 리뉴얼 관리 - 라이선스 만료 추적, 이메일/견적서 자동 발송',
            '2. 고객사 관리 - 고객 정보 통합, 거래 내역 추적',
            '3. 견적서 관리 - 견적서 생성, Status 진행률 관리',
            '4. 매입/매출 관리 - 수익성 분석, 이익율 계산',
            '5. 대시보드 - 월별/팀별 통계, 시각화',
            '6. 인증서 관리 - 5개 제조사 인증서 자동 생성',
            '',
            '• 업무 자동화로 생산성 향상 (이메일/견적서 자동 생성)',
            '• 데이터 기반 의사결정 (대시보드, 통계)',
            '• 실수 방지 (템플릿 기반 자동화)',
            '• 정보 통합 관리 (구글 시트 연동)'
        ]
    }
    
    df = pd.DataFrame(overview_data)
    df.to_excel(writer, sheet_name='1️⃣ 프로젝트 개요', index=False)
    
    worksheet = writer.sheets['1️⃣ 프로젝트 개요']
    worksheet.column_dimensions['A'].width = 20
    worksheet.column_dimensions['B'].width = 80

def create_architecture_sheet(writer):
    """시스템 아키텍처 시트"""
    arch_data = {
        '계층': [
            'Presentation Layer',
            '',
            '',
            '',
            '',
            '',
            '',
            'Business Logic Layer',
            '',
            '',
            '',
            'Data Access Layer',
            '',
            '',
            '',
            'External Integration',
            '',
            ''
        ],
        '컴포넌트': [
            'renewal_gui.py',
            'LoginWindow',
            'MainApp',
            'Sidebar Navigation',
            'Tab Management',
            'Loading Indicator',
            '',
            'renewal_preview.py',
            'customer_info.py',
            'quote_info.py',
            'purchase_sales_info.py',
            'data_loader.py',
            'sendlog.py',
            'organization_manager.py',
            '',
            'Google Sheets API',
            'DAOU 전자결재 (Selenium)',
            'Outlook 이메일 (COM)'
        ],
        '설명': [
            '메인 애플리케이션 - 로그인, 탭 관리',
            '구글 시트 기반 사용자 인증',
            '6개 탭 관리, 데이터 로더 초기화',
            'Material Design 스타일 메뉴',
            'Renewal, 고객사, 견적서, 매입/매출, 대시보드, 인증서',
            '백그라운드 로딩 진행 표시',
            '',
            'Renewal 관리 - 데이터 조회, 이메일/견적서 발송',
            '고객사 관리 - 정보 관리, 거래 내역',
            '견적서 관리 - Status 관리, 견적서 생성',
            '매입/매출 관리 - 수익성 분석',
            '구글 시트 데이터 로딩, 캐싱 (5분 TTL)',
            '로그 기록 및 조회',
            '조직 구조 관리 (부서/팀/영업사원)',
            '',
            '5개 워크시트 연동 (renewal, customer, quote, price, log)',
            'Selenium 기반 자동 결재 요청',
            'win32com 기반 이메일 자동 발송'
        ],
        '기술': [
            'Tkinter, ttk',
            'SHA-256, gspread',
            'Threading',
            'Material Design',
            'Frame, pack()',
            'Toplevel, Progressbar',
            '',
            'pandas, Treeview',
            'pandas, Treeview',
            'pandas, Treeview',
            'pandas, Treeview',
            'gspread, pandas, pickle',
            'gspread',
            'pandas',
            '',
            'gspread, google-auth',
            'selenium, webdriver',
            'pywin32'
        ]
    }
    
    df = pd.DataFrame(arch_data)
    df.to_excel(writer, sheet_name='2️⃣ 시스템 아키텍처', index=False)
    
    worksheet = writer.sheets['2️⃣ 시스템 아키텍처']
    worksheet.column_dimensions['A'].width = 25
    worksheet.column_dimensions['B'].width = 30
    worksheet.column_dimensions['C'].width = 60
    worksheet.column_dimensions['D'].width = 25

def create_modules_sheet(writer):
    """핵심 모듈 시트"""
    modules_data = {
        '탭': [
            '🔄 Renewal 관리',
            '',
            '',
            '',
            '',
            '',
            '👥 고객사 관리',
            '',
            '',
            '',
            '📋 견적서 관리',
            '',
            '',
            '',
            '💰 매입/매출 관리',
            '',
            '',
            '',
            '📊 대시보드',
            '',
            '',
            '',
            '',
            '🔐 인증서 관리',
            '',
            '',
            '',
            ''
        ],
        '모듈': [
            'renewal_preview.py',
            'renewal_email.py',
            'renewal_email_multi.py',
            'renewal_quote.py',
            'renewal_quote_multi.py',
            'renewal_search.py',
            'customer_info.py',
            'customer_add.py',
            'customer_details.py',
            'customer_account_info.py',
            'quote_info.py',
            'quote_new.py',
            'quote_modify.py',
            'quote_status_modify.py',
            'purchase_sales_info.py',
            'purchase_sales_modify.py',
            'purchase_sales_expdate_modify.py',
            '',
            'dashboard_info.py',
            'interactive_dashboard.py',
            'team_performance_dashboard.py',
            'dashboard_foundry.py',
            '',
            'license_certificate_info.py',
            'license_certificate_BorisFX.py',
            'license_certificate_Foundry.py',
            'license_certificate_Marmoset.py',
            'license_certificate_Maxon.py',
            'license_certificate_VideoCopilot.py'
        ],
        '주요 기능': [
            '리뉴얼 데이터 조회, 필터링 (만료일/계산서발행일 기준)',
            '단일 이메일 자동 발송 (Outlook, HTML 템플릿)',
            '다중 선택 이메일 발송',
            '단일 견적서 생성 (Excel -> PDF -> 이메일 첨부)',
            '다중 견적서 생성',
            '검색 및 필터링 유틸리티',
            '고객사 정보 조회, 거래 내역 표시',
            '신규 고객사 등록',
            '고객사 상세 정보 다이얼로그',
            '고객사 계정 정보 관리',
            '견적서 데이터 조회, Status 관리 (10%~100%)',
            '신규 견적서 생성',
            '견적서 수정',
            'Status 일괄 수정',
            '매입/매출 데이터 조회, 수익성 분석',
            '매입/매출 정보 수정',
            '만료일 일괄 수정',
            '',
            '월별 영업사원별 통계, Status 분포 차트',
            '인터랙티브 대시보드, 다양한 차트 타입, 드릴다운',
            'Target 기반 팀별 성과 분석',
            'Foundry 제품 전용 대시보드',
            '',
            '5개 제조사 인증서 생성 메인 UI',
            'BorisFX 인증서 (Word 템플릿)',
            'Foundry 인증서 (Excel + .lic 파일)',
            'Marmoset 인증서 (Excel)',
            'Maxon 인증서 (Excel)',
            'VideoCopilot 인증서 (Excel)'
        ],
        '데이터 소스': [
            'renewal_list',
            'renewal_list',
            'renewal_list',
            'renewal_list + quote 템플릿',
            'renewal_list + quote 템플릿',
            'renewal_list',
            'renewal_list + customer_list',
            'customer_list',
            'renewal_list + customer_list',
            'customer_list',
            'quote_list',
            'quote_list + price_list',
            'quote_list',
            'quote_list',
            'renewal_list',
            'renewal_list',
            'renewal_list',
            '',
            'quote_list',
            'quote_list + renewal_list',
            'Target 워크시트 + quote_list',
            'renewal_list (Foundry만)',
            '',
            'license_certificate/ 템플릿',
            'BorisFX_origin.docx',
            'Foundry_origin.xlsx',
            'Marmoset_origin.xlsx',
            'Maxon_origin.xlsx',
            'VIDEO COPILOT_origin.xlsx'
        ]
    }
    
    df = pd.DataFrame(modules_data)
    df.to_excel(writer, sheet_name='3️⃣ 핵심 모듈', index=False)
    
    worksheet = writer.sheets['3️⃣ 핵심 모듈']
    worksheet.column_dimensions['A'].width = 20
    worksheet.column_dimensions['B'].width = 35
    worksheet.column_dimensions['C'].width = 60
    worksheet.column_dimensions['D'].width = 35

def create_database_sheet(writer):
    """데이터베이스 구조 시트"""
    db_data = {
        '워크시트': [
            'renewal_list',
            '',
            '',
            '',
            '',
            'customer_list',
            '',
            '',
            'quote_list',
            '',
            '',
            '',
            'price_list',
            '',
            '',
            'log_list',
            '',
            ''
        ],
        'Sheet ID': [
            '17rfmG5DEOj1CC-iA6SEVIdsS8zP-UC7lGBT6w2LwBvQ',
            '',
            '',
            '',
            '',
            '1hURXlr1g7qMPoj7GxblVC1SDgM9hCwR6qyG20We1188',
            '',
            '',
            '1xjIjFe1Q9dq2zeQOkBSpbqIUOyXAxbyTO8lKN-r1lw0',
            '',
            '',
            '',
            '1OGt__Olrempc5ZUy6QqqhvgRseFLanv1xCPaw5Y2qJk',
            '',
            '',
            '1llolEKlleT6Cve_M5Qt_CO_JDxQczIKBMqL8fj5vfSk',
            '',
            ''
        ],
        '용도': [
            '리뉴얼 핵심 데이터',
            '- 고객사명, 거래처명, 영업사원',
            '- 제품명, 대분류, 소분류',
            '- 계산서 발행일, 만료일',
            '- 판매 단가, 원가, 이익액, 이익율',
            '고객사 기본 정보',
            '- 고객사명, 담당자명, 직함',
            '- 연락처, 핸드폰, 이메일, 주소',
            '견적서 데이터',
            '- Status (10%~100%), 일자, 영업사원',
            '- 제품, 수량, 판매 단가, 원가',
            '- 고객사명, 거래처명, 만료일',
            '제품별 가격 정보',
            '- 제품명, 대분류, 소분류',
            '- 판매 단가, 원가, 공급업체',
            '작업 로그',
            '- 일시, 고객사명, 제품, 만료일',
            '- 액션, 실행자, 사용자'
        ],
        '접근 방식': [
            'data_loader.get_renewal_data()',
            'renewal_preview.py',
            'customer_info.py',
            'purchase_sales_info.py',
            '',
            'data_loader.get_customer_data()',
            'customer_info.py',
            'customer_add.py',
            'data_loader.get_quote_data()',
            'quote_info.py',
            'dashboard_info.py',
            '',
            'data_loader.get_price_data()',
            'quote_new.py',
            '',
            'sendlog.py',
            'renewal_preview.py',
            ''
        ]
    }
    
    df = pd.DataFrame(db_data)
    df.to_excel(writer, sheet_name='4️⃣ 데이터베이스 구조', index=False)
    
    worksheet = writer.sheets['4️⃣ 데이터베이스 구조']
    worksheet.column_dimensions['A'].width = 20
    worksheet.column_dimensions['B'].width = 50
    worksheet.column_dimensions['C'].width = 45
    worksheet.column_dimensions['D'].width = 30

def create_integration_sheet(writer):
    """외부 시스템 연동 시트"""
    integration_data = {
        '시스템': [
            'Google Sheets API',
            '',
            '',
            '',
            '',
            '',
            'DAOU 전자결재',
            '',
            '',
            '',
            '',
            'Outlook 이메일',
            '',
            '',
            '',
            'Excel 자동화',
            '',
            ''
        ],
        '연동 방식': [
            'gspread + google-auth',
            '서비스 계정 (Service Account)',
            'JSON 키 파일 인증',
            '권한: spreadsheets',
            'API 호출 제한 관리 (429 에러)',
            '재시도 로직 (지수 백오프)',
            'Selenium WebDriver',
            '헤드리스 크롬 브라우저',
            '저장된 크롬 프로필 사용',
            '선택자 관리 (JSON)',
            '로그인 정보 암호화 저장',
            'pywin32 (win32com.client)',
            'Outlook COM 객체',
            'HTML 템플릿 기반',
            '첨부 파일 자동 추가',
            'pywin32 (win32com.client)',
            'Excel COM 객체',
            'PDF 변환 (ExportAsFixedFormat)'
        ],
        '주요 기능': [
            '5개 워크시트 데이터 읽기/쓰기',
            'renewal_list, customer_list, quote_list',
            'price_list, log_list',
            '캐싱 (5분 TTL)',
            '백그라운드 로딩',
            'API 호출 최적화',
            '결재 문서 자동 작성',
            '견적서 첨부 파일 업로드',
            '결재 요청 자동화',
            'UI 선택자 자동 관리',
            '오류 로깅',
            '이메일 자동 작성',
            'HTML 템플릿 (제품별)',
            '템플릿 변수 치환',
            'PDF 견적서 첨부',
            '견적서 템플릿 자동 입력',
            '셀 데이터 쓰기',
            'PDF 변환 및 저장'
        ],
        '설정 파일': [
            'google_sheet_key/renewal-bot-463605-b8f41de8fbdb.json',
            '',
            '',
            '',
            '',
            '',
            'lisa_daou_selectors.json',
            'daou_credentials.json',
            'lisa_daou_sign.log',
            '',
            '',
            'email_temp/*.html',
            'email_template_Chaosgroup.html',
            'email_template_Foundry.html',
            'email_template_TeamViewer.html',
            'quote/quote_origin_01.xlsx',
            '',
            ''
        ]
    }
    
    df = pd.DataFrame(integration_data)
    df.to_excel(writer, sheet_name='5️⃣ 외부 시스템 연동', index=False)
    
    worksheet = writer.sheets['5️⃣ 외부 시스템 연동']
    worksheet.column_dimensions['A'].width = 25
    worksheet.column_dimensions['B'].width = 35
    worksheet.column_dimensions['C'].width = 40
    worksheet.column_dimensions['D'].width = 35

def create_ui_design_sheet(writer):
    """UI/UX 디자인 시트"""
    ui_data = {
        '디자인 요소': [
            '색상 팔레트',
            'Primary',
            'Primary Hover',
            'Secondary',
            'Success',
            'Danger',
            'Warning',
            'Info',
            'Background',
            'Card',
            'Text Primary',
            'Text Secondary',
            'Border',
            '',
            '폰트',
            'Family',
            'Title',
            'Header',
            'Body',
            'Small',
            '',
            '레이아웃',
            'Sidebar Width',
            'Card Padding',
            'Button Height',
            'Input Height'
        ],
        '값/설명': [
            'Material Design 기반',
            '#3B82F6 (파랑)',
            '#2563EB (진한 파랑)',
            '#6B7280 (회색)',
            '#10B981 (초록)',
            '#EF4444 (빨강)',
            '#F59E0B (주황)',
            '#06B6D4 (청록)',
            '#F7F9FB (연한 회색)',
            '#FFFFFF (흰색)',
            '#1F2937 (진한 회색)',
            '#6B7280 (중간 회색)',
            '#E5E7EB (테두리 회색)',
            '',
            '맑은 고딕',
            '맑은 고딕',
            '맑은 고딕, 16, bold',
            '맑은 고딕, 12, bold',
            '맑은 고딕, 10',
            '맑은 고딕, 9',
            '',
            '',
            '200px',
            '20px',
            '40px',
            '32px'
        ],
        '사용처': [
            '',
            '주요 버튼, 선택된 메뉴',
            '버튼 hover 효과',
            '보조 텍스트, 비활성 요소',
            '성공 메시지, 확인 버튼',
            '삭제 버튼, 오류 메시지',
            '경고 메시지',
            '정보 메시지',
            '앱 배경',
            '카드, 입력 필드',
            '제목, 헤더',
            '부가 설명, 힌트',
            '카드 테두리, 구분선',
            '',
            '모든 텍스트',
            '전체 기본 폰트',
            '페이지 제목, 탭 제목',
            '섹션 헤더, 카드 헤더',
            '일반 텍스트, 입력 필드',
            '작은 정보, 힌트',
            '',
            '',
            '좌측 네비게이션 사이드바',
            '카드 내부 여백',
            '버튼 기본 높이',
            '입력 필드 높이'
        ]
    }
    
    df = pd.DataFrame(ui_data)
    df.to_excel(writer, sheet_name='6️⃣ UIUX 디자인', index=False)
    
    worksheet = writer.sheets['6️⃣ UIUX 디자인']
    worksheet.column_dimensions['A'].width = 25
    worksheet.column_dimensions['B'].width = 30
    worksheet.column_dimensions['C'].width = 50

def create_development_guide_sheet(writer):
    """개발 가이드 시트"""
    guide_data = {
        'Step': [
            'Step 1', '', '',
            'Step 2', '', '', '', '', '',
            'Step 3', '', '', '',
            'Step 4', '', '', '',
            'Step 5', '', '',
            'Step 6', '', '',
            'Step 7', '', '',
            'Step 8', '', '',
            'Step 9', '', '',
            'Step 10', '', '',
            'Step 11', '', ''
        ],
        '작업 내용': [
            '프로젝트 초기 설정',
            '- 폴더 구조 생성',
            '- 필요한 라이브러리 설치 (requirements.txt)',
            '메인 애플리케이션 개발',
            '- renewal_gui.py 작성',
            '- LoginWindow 클래스 (구글 시트 로그인)',
            '- MainApp 클래스 (사이드바, 탭 관리)',
            '- ttk.Style 설정 (Material Design)',
            '- 아이콘 리소스 추가',
            '데이터 로더 개발',
            '- data_loader.py 작성',
            '- DataLoader 클래스',
            '- 캐싱 시스템, API 호출 최적화',
            'Renewal 관리 탭 개발',
            '- renewal_preview.py 작성',
            '- 이메일 발송 (renewal_email.py)',
            '- 견적서 생성 (renewal_quote.py)',
            '고객사 관리 탭 개발',
            '- customer_info.py 작성',
            '- 고객사 등록 (customer_add.py)',
            '견적서 관리 탭 개발',
            '- quote_info.py 작성',
            '- Status 관리 (quote_status_modify.py)',
            '매입/매출 관리 탭 개발',
            '- purchase_sales_info.py 작성',
            '- 수익성 분석',
            '대시보드 개발',
            '- dashboard_info.py (기본)',
            '- interactive_dashboard.py (인터랙티브)',
            '인증서 관리 개발',
            '- license_certificate_info.py 작성',
            '- 5개 제조사별 모듈',
            'DAOU 연동 개발',
            '- lisa_daou_sign.py 작성',
            '- Selenium 자동화',
            '패키징 및 배포',
            '- LISA.spec 파일 작성',
            '- PyInstaller 빌드'
        ],
        '예상 소요 시간': [
            '1일', '', '',
            '3일', '', '', '', '', '',
            '2일', '', '', '',
            '5일', '', '', '',
            '3일', '', '',
            '3일', '', '',
            '3일', '', '',
            '5일', '', '',
            '3일', '', '',
            '3일', '', '',
            '1일', '', ''
        ],
        '산출물': [
            '프로젝트 폴더, requirements.txt',
            '',
            '',
            'renewal_gui.py',
            'LoginWindow 클래스',
            'MainApp 클래스',
            '스타일 설정',
            'design/lisa.ico',
            '',
            'data_loader.py',
            'DataLoader 클래스',
            '전역 함수 (캐싱)',
            '',
            'renewal_preview.py',
            'renewal_email.py, renewal_email_multi.py',
            'renewal_quote.py, renewal_quote_multi.py',
            'email_temp/*.html, quote/*.xlsx',
            'customer_info.py',
            'customer_add.py',
            'customer_details.py',
            'quote_info.py',
            'quote_new.py, quote_modify.py',
            'quote_status_modify.py',
            'purchase_sales_info.py',
            'purchase_sales_modify.py',
            'purchase_sales_expdate_modify.py',
            'dashboard_info.py',
            'interactive_dashboard.py',
            'team_performance_dashboard.py',
            'license_certificate_info.py',
            'license_certificate_*.py (5개)',
            'license_certificate/ 템플릿',
            'lisa_daou_sign.py',
            'lisa_daou_selectors.json',
            'daou_credentials.json',
            'LISA.spec',
            'dist/LISA.exe',
            '배포 패키지'
        ]
    }
    
    df = pd.DataFrame(guide_data)
    df.to_excel(writer, sheet_name='7️⃣ 개발 가이드', index=False)
    
    worksheet = writer.sheets['7️⃣ 개발 가이드']
    worksheet.column_dimensions['A'].width = 10
    worksheet.column_dimensions['B'].width = 50
    worksheet.column_dimensions['C'].width = 20
    worksheet.column_dimensions['D'].width = 40

def create_tech_stack_sheet(writer):
    """기술 스택 시트"""
    tech_data = {
        '카테고리': [
            '언어',
            '',
            'GUI',
            '',
            '',
            '데이터 처리',
            '',
            '',
            '구글 API',
            '',
            '시각화',
            '',
            '자동화',
            '',
            '',
            '날짜/시간',
            '',
            '패키징',
            ''
        ],
        '기술/라이브러리': [
            'Python',
            '',
            'Tkinter',
            'ttk',
            'Pillow (PIL)',
            'pandas',
            'openpyxl',
            'numpy',
            'gspread',
            'google-auth',
            'matplotlib',
            '',
            'selenium',
            'webdriver-manager',
            'pywin32',
            'tkcalendar',
            '',
            'PyInstaller',
            ''
        ],
        '버전': [
            '3.x',
            '',
            '내장',
            '내장',
            '>=10.0.0',
            '>=2.0.0',
            '>=3.1.0',
            '>=1.24.0',
            '>=5.10.0',
            '>=2.22.0',
            '>=3.7.0',
            '',
            '>=4.10.0',
            '>=3.9.0',
            '>=306',
            '>=1.6.1',
            '',
            '>=5.13.0',
            ''
        ],
        '용도': [
            '개발 언어',
            '',
            'GUI 프레임워크',
            'Material Design 위젯',
            '이미지 처리, 아이콘 표시',
            '데이터 처리, DataFrame',
            'Excel 파일 읽기/쓰기',
            '수치 계산',
            '구글 시트 API',
            'OAuth2 인증',
            '차트, 그래프 시각화',
            '',
            '웹 자동화 (DAOU)',
            'ChromeDriver 자동 설치',
            'Outlook, Excel COM 객체',
            '달력 위젯',
            '',
            '단일 실행 파일 생성',
            ''
        ]
    }
    
    df = pd.DataFrame(tech_data)
    df.to_excel(writer, sheet_name='8️⃣ 기술 스택', index=False)
    
    worksheet = writer.sheets['8️⃣ 기술 스택']
    worksheet.column_dimensions['A'].width = 20
    worksheet.column_dimensions['B'].width = 30
    worksheet.column_dimensions['C'].width = 20
    worksheet.column_dimensions['D'].width = 50

def create_deployment_sheet(writer):
    """배포 가이드 시트"""
    deploy_data = {
        '단계': [
            '1. 사전 준비',
            '',
            '',
            '',
            '2. LISA.spec 작성',
            '',
            '',
            '',
            '',
            '',
            '3. 빌드 실행',
            '',
            '',
            '4. 테스트',
            '',
            '',
            '5. 배포 패키지 구성',
            '',
            '',
            '',
            '6. 배포',
            '',
            ''
        ],
        '작업 내용': [
            'PyInstaller 설치',
            '모든 리소스 파일 확인',
            '아이콘 파일 (design/lisa.ico)',
            'VERSION.txt 파일 업데이트',
            'LISA.spec 파일 작성',
            '- 진입점: renewal_gui.py',
            '- 단일 파일: onefile',
            '- 윈도우 모드: noconsole',
            '- 아이콘: design/lisa.ico',
            '- 추가 데이터: design, google_sheet_key, templates 등',
            'PyInstaller 명령어 실행',
            'pyinstaller LISA.spec',
            '빌드 완료 확인 (dist/LISA.exe)',
            '실행 파일 테스트',
            '- 로그인 기능',
            '- 각 탭 동작 확인',
            '배포 폴더 구성',
            '- LISA.exe',
            '- VERSION.txt',
            '- README.txt (사용 설명서)',
            '배포 위치로 복사',
            '사용자에게 안내',
            '업데이트 공지'
        ],
        '명령어/파일': [
            'pip install pyinstaller',
            '',
            'design/lisa.ico',
            'VERSION.txt',
            'LISA.spec',
            '',
            '',
            '',
            '',
            '',
            'pyinstaller LISA.spec',
            '또는',
            'pyinstaller --icon=design/lisa.ico --noconsole --onefile --name LISA renewal_gui.py',
            'dist/LISA.exe',
            '',
            '',
            'dist/',
            'LISA.exe',
            'VERSION.txt',
            'README.txt',
            '',
            '',
            ''
        ],
        '주의사항': [
            'Python 3.x 환경',
            '모든 리소스 파일 포함',
            '',
            'v2.01 확인',
            'add-data 정확히 지정',
            '',
            '',
            '',
            '',
            '빌드 오류 확인',
            '빌드 시간: 약 5-10분',
            'build/, dist/ 폴더 생성됨',
            '',
            '다른 PC에서도 테스트',
            '구글 시트 API 키 확인',
            '외부 의존성 확인',
            '압축 파일로 배포',
            '실행 파일만 배포 가능',
            '',
            '',
            '백업 권장',
            '변경사항 문서화',
            '버전 히스토리 관리'
        ]
    }
    
    df = pd.DataFrame(deploy_data)
    df.to_excel(writer, sheet_name='9️⃣ 배포 가이드', index=False)
    
    worksheet = writer.sheets['9️⃣ 배포 가이드']
    worksheet.column_dimensions['A'].width = 20
    worksheet.column_dimensions['B'].width = 50
    worksheet.column_dimensions['C'].width = 60
    worksheet.column_dimensions['D'].width = 40

def create_checklist_sheet(writer):
    """체크리스트 시트"""
    checklist_data = {
        '구분': [
            '개발 완료',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '테스트',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '문서화',
            '',
            '',
            ''
        ],
        '항목': [
            '로그인 시스템',
            '메인 애플리케이션 (사이드바 + 탭)',
            '데이터 로더 (구글 시트 연동)',
            'Renewal 관리 탭',
            '고객사 관리 탭',
            '견적서 관리 탭',
            '매입/매출 관리 탭',
            '대시보드 (기본)',
            '대시보드 (인터랙티브)',
            '대시보드 (팀별)',
            '인증서 관리 (BorisFX)',
            '인증서 관리 (Foundry)',
            '인증서 관리 (Marmoset, Maxon, VideoCopilot)',
            '이메일 발송 (Outlook)',
            '견적서 생성 (Excel + PDF)',
            'DAOU 전자결재 연동',
            '로그 관리',
            '로딩 인디케이터',
            'PyInstaller 패키징',
            '로그인 기능 테스트',
            '데이터 로딩 및 캐싱 테스트',
            '필터링 및 검색 테스트',
            '이메일 발송 테스트',
            '견적서 생성 테스트',
            'DAOU 연동 테스트',
            '인증서 생성 테스트',
            '대시보드 차트 테스트',
            '로그 기록 및 조회 테스트',
            '사용자 매뉴얼 작성',
            'API 문서 작성',
            '릴리즈 노트 작성',
            'README 파일 작성'
        ],
        '상태': [
            '✅ 완료',
            '✅ 완료',
            '✅ 완료',
            '✅ 완료',
            '✅ 완료',
            '✅ 완료',
            '✅ 완료',
            '✅ 완료',
            '✅ 완료',
            '✅ 완료',
            '✅ 완료',
            '✅ 완료',
            '✅ 완료',
            '✅ 완료',
            '✅ 완료',
            '✅ 완료',
            '✅ 완료',
            '✅ 완료',
            '✅ 완료',
            '✅ 테스트 완료',
            '✅ 테스트 완료',
            '✅ 테스트 완료',
            '✅ 테스트 완료',
            '✅ 테스트 완료',
            '✅ 테스트 완료',
            '✅ 테스트 완료',
            '✅ 테스트 완료',
            '✅ 테스트 완료',
            '✅ 완료',
            '📝 작성 중',
            '✅ 완료',
            '✅ 완료'
        ],
        '담당자': [
            '개발팀',
            '개발팀',
            '개발팀',
            '개발팀',
            '개발팀',
            '개발팀',
            '개발팀',
            '개발팀',
            '개발팀',
            '개발팀',
            '개발팀',
            '개발팀',
            '개발팀',
            '개발팀',
            '개발팀',
            '개발팀',
            '개발팀',
            '개발팀',
            '개발팀',
            'QA팀',
            'QA팀',
            'QA팀',
            'QA팀',
            'QA팀',
            'QA팀',
            'QA팀',
            'QA팀',
            'QA팀',
            '기획팀',
            '개발팀',
            '기획팀',
            '기획팀'
        ],
        '비고': [
            'SHA-256 해시, 구글 시트 인증',
            'Material Design, 6개 탭',
            '5분 캐싱, 백그라운드 로딩',
            '이메일/견적서 발송 포함',
            '신규 등록, 상세 정보',
            'Status 관리 (10%~100%)',
            '수익성 분석',
            '월별 영업사원별 통계',
            '다양한 차트, 드릴다운',
            'Target 기반 성과 분석',
            'Word 템플릿',
            'Excel + .lic 파일',
            'Excel 템플릿',
            'HTML 템플릿, COM 객체',
            'Excel -> PDF 변환',
            'Selenium 자동화',
            '구글 시트 log_list',
            'Toplevel + Progressbar',
            '단일 실행 파일 (onefile)',
            '로그인/로그아웃',
            '초기 로딩, 전체 로딩',
            '필터 동작, 검색 기능',
            '단일/다중 발송',
            '단일/다중 생성',
            '자동 로그인, 결재 요청',
            '5개 제조사 모두',
            'matplotlib 차트',
            '로그 조회 (기간별)',
            '설치/사용법',
            'API 키, 구조',
            '버전별 변경사항',
            '프로젝트 소개'
        ]
    }
    
    df = pd.DataFrame(checklist_data)
    df.to_excel(writer, sheet_name='🔟 체크리스트', index=False)
    
    worksheet = writer.sheets['🔟 체크리스트']
    worksheet.column_dimensions['A'].width = 15
    worksheet.column_dimensions['B'].width = 50
    worksheet.column_dimensions['C'].width = 20
    worksheet.column_dimensions['D'].width = 15
    worksheet.column_dimensions['E'].width = 40

def main():
    """메인 실행 함수"""
    print("="*80)
    print("📋 LISA 프로젝트 기획서 생성 도구")
    print("="*80)
    print()
    
    md_path = 'LISA_재현_프롬프트.md'
    output_path = 'LISA_프로젝트_기획서.xlsx'
    
    try:
        create_excel_planning_doc(md_path, output_path)
        
        print()
        print("="*80)
        print(f"✅ 완료! Excel 기획서가 생성되었습니다: {output_path}")
        print()
        print("📊 생성된 시트:")
        print("   1. 📋 표지")
        print("   2. 📑 목차")
        print("   3. 1️⃣ 프로젝트 개요")
        print("   4. 2️⃣ 시스템 아키텍처")
        print("   5. 3️⃣ 핵심 모듈")
        print("   6. 4️⃣ 데이터베이스 구조")
        print("   7. 5️⃣ 외부 시스템 연동")
        print("   8. 6️⃣ UI/UX 디자인")
        print("   9. 7️⃣ 개발 가이드")
        print("  10. 8️⃣ 기술 스택")
        print("  11. 9️⃣ 배포 가이드")
        print("  12. 🔟 체크리스트")
        print("="*80)
        
    except Exception as e:
        print(f"❌ 오류 발생: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()

