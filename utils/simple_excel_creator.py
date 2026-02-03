# -*- coding: utf-8 -*-
import pandas as pd
from datetime import datetime
import os

# 현재 스크립트의 디렉토리를 작업 디렉토리로 설정
os.chdir(os.path.dirname(os.path.abspath(__file__)))

output_path = 'LISA_프로젝트_기획서.xlsx'

print(f"📊 Excel 기획서 생성 중: {output_path}")

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    
    # 1. 표지
    cover = pd.DataFrame({
        '항목': ['프로젝트명', '버전', '작성일', '상태'],
        '내용': ['LISA (License Information System for Administration)', 
                'v2.01', 
                datetime.now().strftime('%Y-%m-%d'), 
                '운영 중']
    })
    cover.to_excel(writer, sheet_name='📋 표지', index=False)
    
    # 2. 프로젝트 개요
    overview = pd.DataFrame({
        '구분': ['프로젝트명', '영문명', '버전', '개발 언어', 'UI 프레임워크', '목적'],
        '내용': [
            'LISA',
            'License Information System for Administration',
            'v2.01',
            'Python 3.x',
            'Tkinter (Material Design)',
            '소프트웨어 라이선스 관리 및 영업 지원 자동화'
        ]
    })
    overview.to_excel(writer, sheet_name='1️⃣ 프로젝트 개요', index=False)
    
    # 3. 핵심 모듈
    modules = pd.DataFrame({
        '탭': ['Renewal 관리', '고객사 관리', '견적서 관리', '매입/매출 관리', '대시보드', '인증서 관리'],
        '주요 모듈': [
            'renewal_preview.py, renewal_email.py, renewal_quote.py',
            'customer_info.py, customer_add.py, customer_details.py',
            'quote_info.py, quote_new.py, quote_modify.py',
            'purchase_sales_info.py, purchase_sales_modify.py',
            'dashboard_info.py, interactive_dashboard.py, team_performance_dashboard.py',
            'license_certificate_info.py, license_certificate_*.py'
        ],
        '주요 기능': [
            '리뉴얼 데이터 조회, 이메일/견적서 자동 발송, DAOU 연동',
            '고객사 정보 관리, 거래 내역 추적, 신규 등록',
            '견적서 생성/수정, Status 관리 (10%~100%)',
            '매입/매출 분석, 수익성 분석, 이익율 계산',
            '월별/팀별 통계, 차트 시각화, Target 기반 성과 분석',
            '5개 제조사 인증서 자동 생성 (BorisFX, Foundry, Marmoset, Maxon, VideoCopilot)'
        ]
    })
    modules.to_excel(writer, sheet_name='2️⃣ 핵심 모듈', index=False)
    
    # 4. 데이터베이스 구조
    database = pd.DataFrame({
        '워크시트': ['renewal_list', 'customer_list', 'quote_list', 'price_list', 'log_list'],
        'Sheet ID': [
            '17rfmG5DEOj1CC-iA6SEVIdsS8zP-UC7lGBT6w2LwBvQ',
            '1hURXlr1g7qMPoj7GxblVC1SDgM9hCwR6qyG20We1188',
            '1xjIjFe1Q9dq2zeQOkBSpbqIUOyXAxbyTO8lKN-r1lw0',
            '1OGt__Olrempc5ZUy6QqqhvgRseFLanv1xCPaw5Y2qJk',
            '1llolEKlleT6Cve_M5Qt_CO_JDxQczIKBMqL8fj5vfSk'
        ],
        '용도': [
            '리뉴얼 핵심 데이터 (고객사명, 제품, 만료일, 판매 단가, 원가 등)',
            '고객사 기본 정보 (담당자, 연락처, 이메일, 주소 등)',
            '견적서 데이터 (Status, 일자, 영업사원, 제품, 판매 합계 등)',
            '제품별 가격 정보 (제품명, 판매 단가, 원가 등)',
            '작업 로그 (일시, 고객사명, 제품, 액션, 실행자 등)'
        ]
    })
    database.to_excel(writer, sheet_name='3️⃣ 데이터베이스 구조', index=False)
    
    # 5. 외부 시스템 연동
    integration = pd.DataFrame({
        '시스템': ['Google Sheets API', 'DAOU 전자결재', 'Outlook 이메일', 'Excel 자동화'],
        '연동 방식': [
            'gspread + google-auth, 서비스 계정 인증',
            'Selenium WebDriver, 헤드리스 크롬',
            'pywin32 (win32com.client), COM 객체',
            'pywin32 (win32com.client), PDF 변환'
        ],
        '주요 기능': [
            '5개 워크시트 읽기/쓰기, 캐싱 (5분 TTL)',
            '결재 문서 자동 작성, 첨부 파일 업로드',
            '이메일 자동 작성, HTML 템플릿, PDF 첨부',
            '견적서 템플릿 자동 입력, PDF 변환'
        ]
    })
    integration.to_excel(writer, sheet_name='4️⃣ 외부 시스템 연동', index=False)
    
    # 6. 기술 스택
    tech_stack = pd.DataFrame({
        '카테고리': ['언어', 'GUI', '데이터 처리', '구글 API', '시각화', '자동화', '패키징'],
        '기술/라이브러리': [
            'Python 3.x',
            'Tkinter, ttk, Pillow',
            'pandas, openpyxl, numpy',
            'gspread, google-auth',
            'matplotlib',
            'selenium, pywin32',
            'PyInstaller'
        ],
        '용도': [
            '개발 언어',
            'GUI 프레임워크, Material Design',
            '데이터 처리, Excel 파일',
            '구글 시트 API, OAuth2 인증',
            '차트, 그래프 시각화',
            '웹 자동화, Outlook/Excel COM',
            '단일 실행 파일 생성'
        ]
    })
    tech_stack.to_excel(writer, sheet_name='5️⃣ 기술 스택', index=False)
    
    # 7. 개발 가이드
    guide = pd.DataFrame({
        'Step': ['Step 1', 'Step 2', 'Step 3', 'Step 4', 'Step 5', 'Step 6'],
        '작업 내용': [
            '프로젝트 초기 설정 - 폴더 구조, requirements.txt',
            '메인 애플리케이션 개발 - renewal_gui.py, 로그인, 탭 관리',
            '데이터 로더 개발 - data_loader.py, 캐싱, API 최적화',
            'Renewal 관리 탭 - renewal_preview.py, 이메일/견적서 발송',
            '나머지 5개 탭 개발 - 고객사, 견적서, 매입/매출, 대시보드, 인증서',
            'DAOU 연동, 패키징 - lisa_daou_sign.py, PyInstaller'
        ],
        '예상 소요 시간': ['1일', '3일', '2일', '5일', '15일', '4일']
    })
    guide.to_excel(writer, sheet_name='6️⃣ 개발 가이드', index=False)
    
    # 8. 배포 가이드
    deployment = pd.DataFrame({
        '단계': ['1. 사전 준비', '2. 빌드 실행', '3. 테스트', '4. 배포'],
        '작업 내용': [
            'PyInstaller 설치, 리소스 파일 확인, VERSION.txt 업데이트',
            'pyinstaller LISA.spec 실행, dist/LISA.exe 생성 확인',
            '실행 파일 테스트 - 로그인, 각 탭 동작 확인',
            '배포 폴더 구성, 사용자 안내, 업데이트 공지'
        ]
    })
    deployment.to_excel(writer, sheet_name='7️⃣ 배포 가이드', index=False)
    
    # 9. 체크리스트
    checklist = pd.DataFrame({
        '구분': ['개발'] * 6 + ['테스트'] * 6,
        '항목': [
            '로그인 시스템', 'Renewal 관리', '고객사 관리', '견적서 관리', '매입/매출 관리', '대시보드',
            '로그인 테스트', '이메일 발송 테스트', '견적서 생성 테스트', 'DAOU 연동 테스트', '인증서 생성 테스트', '대시보드 차트 테스트'
        ],
        '상태': ['✅ 완료'] * 12
    })
    checklist.to_excel(writer, sheet_name='8️⃣ 체크리스트', index=False)
    
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
            adjusted_width = min(max_length + 2, 100)
            worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

print(f"✅ Excel 기획서 생성 완료: {output_path}")
print()
print("📊 생성된 시트:")
print("   1. 📋 표지")
print("   2. 1️⃣ 프로젝트 개요")
print("   3. 2️⃣ 핵심 모듈")
print("   4. 3️⃣ 데이터베이스 구조")
print("   5. 4️⃣ 외부 시스템 연동")
print("   6. 5️⃣ 기술 스택")
print("   7. 6️⃣ 개발 가이드")
print("   8. 7️⃣ 배포 가이드")
print("   9. 8️⃣ 체크리스트")

