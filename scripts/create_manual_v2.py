# -*- coding: utf-8 -*-
"""
LISA 상세 매뉴얼 - 간단 버전 (오류 없이 빠르게 생성)
"""
import pandas as pd
from datetime import datetime
import os

os.chdir(os.path.dirname(os.path.abspath(__file__)))
output_path = 'LISA_상세매뉴얼_UI기획서.xlsx'

print(f"📚 LISA 상세 매뉴얼 생성 중: {output_path}")

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    
    # 1. 표지
    pd.DataFrame({
        '항목': ['문서명', '프로젝트', '버전', '작성일'],
        '내용': ['LISA 상세 매뉴얼 & UI 기획서', 'LISA v2.01', 'v1.0', datetime.now().strftime('%Y-%m-%d')]
    }).to_excel(writer, sheet_name='📋 표지', index=False)
    
    # 2. 메인 애플리케이션 (renewal_gui.py)
    pd.DataFrame({
        'UI 구성': ['로그인 창', 'ID/PW 입력', '로그인 버튼', '', '사이드바 메뉴', 'Renewal 관리', '고객사 관리', '견적서', '매입/매출', '대시보드', '인증서', '', '메인 프레임', '로딩 인디케이터'],
        '설명': ['350x250 다이얼로그', 'ID/PW 입력 필드', '구글 시트 인증 (SHA-256)', '', '좌측 네비게이션 (200px)', '리뉴얼 데이터 관리', '고객사 정보 관리', '견적서 Status 관리', '매입/매출 분석', '통계 차트 (3종)', '인증서 자동 생성', '', '탭 컨텐츠 영역', '백그라운드 로딩 표시'],
        '사용 방법': ['LISA 실행 시 표시', 'ID/PW 입력 후 로그인', 'Enter 또는 클릭', '', '메뉴 클릭으로 탭 전환', '만료일 기준 데이터 조회', '고객사별 거래 내역', 'Status 10%~100% 관리', '수익성 분석', '월별/팀별 통계', '5개 제조사 인증서', '', '선택된 탭 표시', '데이터 로딩 대기']
    }).to_excel(writer, sheet_name='1️⃣ 메인앱 (renewal_gui)', index=False)
    
    # 3. Renewal 관리 (renewal_preview.py)
    pd.DataFrame({
        'UI 요소': ['기준 선택', '만료일/계산서발행일', '', '기간 선택', '이번달', '다음달', '이번분기/다음분기', '사용자정의', '', '필터', '고객사명', '거래처명', '대분류', '영업사원', '', '데이터 테이블', '정렬 가능', '다중 선택', '색상 코딩', '', '액션 버튼', '이메일 작성', '견적서+메일', '여러 제품 선택', '로그 조회'],
        '기능': ['라디오 버튼', '기준 변경 시 재조회', '', '버튼 클릭', 'load_this_month()', 'load_next_month()', 'load_quarter()', 'DateEntry 달력', '', '드롭다운 필터', '고객사 선택', '거래처 선택', '제품 분류', '담당자 선택', '', 'Treeview 위젯', '헤더 클릭 정렬', 'Ctrl/Shift 선택', '만료 임박 빨간색', '', '선택 행 액션', 'Outlook 이메일', 'Excel 견적서 + PDF', '다중 이메일/견적서', 'show_logs()'],
        '사용법': ['1. 기준 선택 (만료일 권장)', '', '', '2. 기간 선택', '현재 월 데이터', '다음 월 데이터', '분기 데이터', '시작~종료 날짜', '', '3. 필요시 필터 적용', '특정 고객사', '특정 거래처', '특정 제품군', '특정 담당자', '', '4. 데이터 확인', '클릭하여 정렬', '여러 행 선택', '빨간색 = 만료 임박', '', '5. 액션 실행', '이메일 자동 작성', '견적서 자동 생성', '복수 선택 처리', '기간별 로그 확인']
    }).to_excel(writer, sheet_name='2️⃣ Renewal 관리', index=False)
    
    # 4. 고객사 관리 (customer_info.py)
    pd.DataFrame({
        '기능': ['고객사 정보 조회', '거래 내역 표시', '통계 표시', '신규 고객사 등록', '고객사 정보 수정', '상세 정보 보기'],
        '설명': ['renewal_list + customer_list 통합', '고객사별 제품 거래 내역', '총 거래건수, 판매액, 이익액', 'CustomerAddWindow 다이얼로그', '구글 시트 직접 업데이트', 'show_customer_detail()'],
        'UI': ['Treeview 테이블', '고객사명, 제품, 판매액 컬럼', '상단 통계 카드', '신규 등록 버튼', '행 더블클릭 또는 버튼', '우클릭 메뉴'],
        '사용법': ['고객사 관리 탭 클릭', '기간 선택 후 필터 적용', '통계 자동 계산 표시', '버튼 클릭 > 정보 입력', '행 선택 > 수정 버튼', '우클릭 > 상세 정보']
    }).to_excel(writer, sheet_name='3️⃣ 고객사 관리', index=False)
    
    # 5. 견적서 관리 (quote_info.py)
    pd.DataFrame({
        '기능': ['Status 관리', 'Status 표시', 'Status 수정', '견적서 생성', '견적서 수정', 'DAOU 연동'],
        '설명': ['10%, 20%, ..., 100% 진행률', '색상 코딩 (회색/노란색/초록색)', '우클릭으로 Status 변경', 'QuoteNewWindow', 'QuoteModifyWindow', 'Selenium 자동화'],
        'Status 의미': ['10~70%: 진행 중 (회색)', '80~90%: 거의 완료 (노란색)', '100%: 계약 완료 (초록색)', '', '', ''],
        '사용법': ['견적서 탭 선택', 'Status별 색상 확인', '우클릭 > Status 변경', '신규 버튼 > 정보 입력', '행 선택 > 수정 버튼', '우클릭 > DAOU 연동']
    }).to_excel(writer, sheet_name='4️⃣ 견적서 관리', index=False)
    
    # 6. 매입/매출 관리 (purchase_sales_info.py)
    pd.DataFrame({
        '분석 항목': ['판매 단가', '원가', '이익액', '이익율', '실 매입 단가', '실 매입 합계', '실 이익액', '실 이익율'],
        '설명': ['고객에게 판매한 가격', '예상 매입 가격', '판매가 - 원가', '(이익액/판매가) × 100', '실제 매입한 가격', '실제 매입 총액', '실제 이익', '실제 이익율'],
        '통계': ['총 판매액', '총 원가', '총 이익액', '평균 이익율', '', '', '', ''],
        '용도': ['매출 확인', '비용 확인', '수익 확인', '수익성 분석', '실제 매입가 확인', '실제 비용 확인', '실제 수익 확인', '실제 수익성 분석']
    }).to_excel(writer, sheet_name='5️⃣ 매입매출 관리', index=False)
    
    # 7. 대시보드 (3종)
    pd.DataFrame({
        '대시보드': ['기본 대시보드', 'dashboard_info.py', '', '인터랙티브 대시보드', 'interactive_dashboard.py', '', '팀별 성과 대시보드', 'team_performance_dashboard.py'],
        '기능': ['월별 영업사원별 통계', 'Status 분포 파이차트', '영업사원별 매출 막대차트', '다양한 차트 타입', '조직 필터 (부서/팀/영업사원)', '드릴다운 기능', 'Target 기반 성과 분석', 'Target vs 실적 비교'],
        '데이터 소스': ['quote_list', '', '', 'quote_list + renewal_list', '', '', 'Target 워크시트', 'quote_list'],
        '사용법': ['대시보드 탭 > 기본', '월 선택 > 차트 확인', '', '대시보드 탭 > 인터랙티브', '차트 타입 선택 > 조직 필터', '차트 클릭하여 상세 보기', '대시보드 탭 > 팀별', '연도/월 선택 > 성과 확인']
    }).to_excel(writer, sheet_name='6️⃣ 대시보드 (3종)', index=False)
    
    # 8. 인증서 관리 (license_certificate_info.py)
    pd.DataFrame({
        '제조사': ['BorisFX', 'Foundry', 'Marmoset', 'Maxon', 'VideoCopilot'],
        '템플릿 형식': ['Word (.docx)', 'Excel (.xlsx) + .lic', 'Excel (.xlsx)', 'Excel (.xlsx)', 'Excel (.xlsx)'],
        '생성 파일': ['PDF', 'Excel + PDF + .lic', 'Excel', 'Excel', 'Excel'],
        '버튼 색상': ['보라색 #8B5CF6', '파란색 #3B82F6', '초록색 #10B981', '주황색 #F59E0B', '빨간색 #EF4444'],
        '사용법': ['버튼 클릭', '고객사 정보 입력', '생성 버튼', '자동으로 템플릿 채우기', 'PDF 변환 (선택)']
    }).to_excel(writer, sheet_name='7️⃣ 인증서 관리', index=False)
    
    # 9. 이메일 발송 (renewal_email.py)
    pd.DataFrame({
        '함수': ['send_emails()', 'send_multi_emails()', '', '', '', '', '', ''],
        '기능': ['단일 제품 이메일', '다중 제품 이메일', '', '', '', '', '', ''],
        'HTML 템플릿': ['email_template.html', 'email_template_Chaosgroup.html', 'email_template_Foundry.html', 'email_template_TeamViewer.html', '', '', '', ''],
        '템플릿 변수': ['{{고객사명}}', '{{제품}}', '{{수량}}', '{{만료일}}', '{{담당자명}}', '{{연락처}}', '{{이메일}}', '{{주소}}'],
        '프로세스': ['1. Outlook COM 객체 생성', '2. HTML 템플릿 로드', '3. 변수 치환', '4. 이메일 창 표시', '5. 검토 후 전송', '6. 로그 자동 기록', '', '']
    }).to_excel(writer, sheet_name='8️⃣ 이메일 발송', index=False)
    
    # 10. 견적서 생성 (renewal_quote.py)
    pd.DataFrame({
        '함수': ['create_and_send_quote()', 'create_multi_quotes()', '', '', '', ''],
        '기능': ['단일 견적서 생성', '다중 견적서 생성', '', '', '', ''],
        '템플릿': ['quote/quote_origin_01.xlsx', '', '', '', '', ''],
        '프로세스': ['1. Excel COM 객체 생성', '2. 템플릿 파일 열기', '3. 셀 데이터 입력', '4. PDF 변환', '5. Outlook에 첨부', '6. 이메일 창 표시'],
        '셀 위치': ['E3: 일자', 'C4: 고객사명', 'E6: 담당자명', 'E7: 연락처', 'A10~: 제품 목록', '']
    }).to_excel(writer, sheet_name='9️⃣ 견적서 생성', index=False)
    
    # 11. DAOU 전자결재 (lisa_daou_sign.py)
    pd.DataFrame({
        '단계': ['1단계', '2단계', '3단계', '4단계', '5단계', '6단계'],
        '작업': ['Selenium 초기화', 'DAOU 로그인', '결재 문서 작성', '첨부 파일 업로드', '결재선 선택', '결재 요청 제출'],
        '설명': ['ChromeDriver 설정', '저장된 프로필 사용', '제목/내용 입력', '견적서 PDF 등', '자동 선택', '자동 클릭'],
        '설정 파일': ['', 'daou_credentials.json', 'lisa_daou_selectors.json', '', '', ''],
        '사용법': ['우클릭 > DAOU 연동', '자동 로그인 확인', '자동 작성 대기', '자동 업로드 대기', '자동 처리', '완료 확인']
    }).to_excel(writer, sheet_name='🔟 DAOU 전자결재', index=False)
    
    # 12. 데이터 로더 (data_loader.py)
    pd.DataFrame({
        '기능': ['구글 시트 연동', '캐싱 시스템', 'API 최적화', '백그라운드 로딩', ''],
        '워크시트': ['renewal_list', 'customer_list', 'quote_list', 'price_list', 'log_list'],
        '캐시 전략': ['메모리 캐싱 5분 TTL', '초기 로딩: 현재 월만', '백그라운드: 전체 데이터', 'API 호출 최소화', ''],
        '메서드': ['load_current_month_data()', 'load_all_data()', 'get_renewal_data()', 'get_customer_data()', 'get_quote_data()']
    }).to_excel(writer, sheet_name='1️⃣1️⃣ 데이터 로더', index=False)
    
    # 13. 사용자 시나리오
    pd.DataFrame({
        '시나리오': ['📧 리뉴얼 이메일 발송', '', '', '📋 견적서 생성 발송', '', '', '👥 고객사 정보 조회', '', '💰 수익성 분석', '', '📊 월별 실적 확인', '', '🔐 인증서 발급', ''],
        '단계': ['1. Renewal 관리 탭', '2. 만료일 기준 선택', '3. 이번달 데이터 조회', '4. 고객 선택 > 이메일 작성', '', '', '1. Renewal 관리 탭', '2. 고객 선택 > 견적서+메일', '3. Excel 자동 생성 > PDF 변환', '', '1. 고객사 관리 탭', '2. 기간/필터 설정', '3. 거래 내역 확인', ''],
        '소요 시간': ['총 5분', '10초', '10초', '1분', '2분 (검토)', '', '총 7분', '30초', '1분', '3분 (검토)', '총 3분', '10초', '2분', ''],
        '결과': ['이메일 창 표시', '템플릿 자동 채움', '검토 후 전송', '로그 자동 기록', '', '', '견적서 PDF 생성', '이메일 자동 첨부', '검토 후 전송', '', '고객사 통계 확인', '거래 내역 분석', '', '']
    }).to_excel(writer, sheet_name='🎬 사용자 시나리오', index=False)
    
    # 14. 문제 해결
    pd.DataFrame({
        '문제': ['로그인 실패', '데이터가 안 보임', 'Outlook 오류', 'Excel 오류', 'DAOU 연동 실패', '느린 로딩'],
        '원인': ['ID/PW 불일치, 구글 시트 권한', '캐시 문제, API 오류', 'Outlook 미실행, COM 오류', 'Excel 미실행, 템플릿 없음', '로그인 세션 만료, 선택자 변경', '전체 데이터 로딩 중, API 제한'],
        '해결방법': ['1. ID/PW 확인\n2. 구글 시트 공유 확인', '1. 새로고침 (F5)\n2. LISA 재시작', '1. Outlook 먼저 실행\n2. Outlook 재시작', '1. Excel 먼저 실행\n2. 템플릿 파일 확인', '1. Chrome에서 DAOU 로그인\n2. 선택자 JSON 확인', '1. 초기 로딩 대기\n2. 백그라운드 완료 대기']
    }).to_excel(writer, sheet_name='🔧 문제 해결', index=False)
    
    # 15. 빠른 참조
    pd.DataFrame({
        '항목': ['프로젝트명', '버전', '메인 파일', '데이터베이스', '', '', '', '', '주요 기능', '', '', '', '', '', '단축키', '', '', ''],
        '내용': ['LISA (License Information System for Administration)', 'v2.01', 'renewal_gui.py', '구글 시트 5개:', '  - renewal_list (리뉴얼)', '  - customer_list (고객사)', '  - quote_list (견적서)', '  - price_list (가격)', '6개 탭:', '  1. Renewal 관리', '  2. 고객사 관리', '  3. 견적서 관리', '  4. 매입/매출 관리', '  5. 대시보드 (3종)', 'F5: 새로고침', 'Ctrl+클릭: 다중 선택', 'Shift+클릭: 범위 선택', '우클릭: 상세 메뉴']
    }).to_excel(writer, sheet_name='📌 빠른 참조', index=False)
    
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

print(f"✅ 완료! {output_path}")
print("\n📚 생성된 시트:")
sheets = ['📋 표지', '1️⃣ 메인앱', '2️⃣ Renewal 관리', '3️⃣ 고객사 관리', 
          '4️⃣ 견적서 관리', '5️⃣ 매입매출 관리', '6️⃣ 대시보드', '7️⃣ 인증서 관리',
          '8️⃣ 이메일 발송', '9️⃣ 견적서 생성', '🔟 DAOU 전자결재', '1️⃣1️⃣ 데이터 로더',
          '🎬 사용자 시나리오', '🔧 문제 해결', '📌 빠른 참조']
for i, sheet in enumerate(sheets, 1):
    print(f"   {i:2d}. {sheet}")

