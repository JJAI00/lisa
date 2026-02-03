import os
import calendar
import tkinter as tk
from tkinter import messagebox, ttk
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

from . import renewal_search as rs
from email_module import renewal_email as res_email
from . import renewal_quote as res_quote
from . import renewal_quote_multi as res_quote_multi
from email_module.renewal_email_multi import send_multi_emails
from core.sendlog import show_logs, KEY_FILE, SCOPES, LOG_SHEET_ID, LOG_SHEET_NAME
from modules.customer.customer_details import show_customer_detail
from integrations.lisa_logging import setup_logger

# (전역 폰트/스타일 블록 삭제됨)

def safe_int_convert(value, default=0):
    """안전한 int 변환 함수"""
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

class RenewalMailerGUI:
    """메인 GUI (미리보기 + 액션)"""
    def __init__(self, root, username=None, data_loader=None):
        self.root     = root
        self.username = username or '테스트유저'
        self.data_loader = data_loader
        if hasattr(root, 'title'):
            root.title(f"LISA v1.5 - {self.username}")
        if hasattr(root, 'geometry'):
            root.geometry("1200x650")

        # 기준 선택 변수 (기본값: 만료일)
        self.criteria_var = tk.StringVar(value="만료일")

        # 데이터 로더가 있으면 사용, 없으면 기존 방식 사용
        if self.data_loader:
            try:
                self.raw_df = self.data_loader.get_renewal_data()
                self.log_df = self.data_loader.get_log_data()
                
                # 데이터가 비어있으면 기존 방식으로 로드
                if self.raw_df.empty:
                    print("Renewal 관리: 데이터 로더에서 데이터를 가져올 수 없어 기존 방식으로 로드합니다.")
                    self.raw_df = self._load_data_with_retry()
                    self.log_df = pd.DataFrame(columns=['고객사명','제품','만료일','액션','일시'])
                    self._load_logs()
                else:
                    print(f"Renewal 관리: 데이터 로더에서 데이터를 성공적으로 가져왔습니다: {len(self.raw_df)}행")
            except Exception as e:
                print(f"Renewal 관리: 데이터 로더 오류 - {e}, 기존 방식으로 로드합니다.")
                self.raw_df = self._load_data_with_retry()
                self.log_df = pd.DataFrame(columns=['고객사명','제품','만료일','액션','일시'])
                self._load_logs()
        else:
            # 기존 방식 (하위 호환성)
            print("Renewal 관리: 기존 방식으로 데이터를 로드합니다.")
            self.raw_df = self._load_data_with_retry()
            self.log_df = pd.DataFrame(columns=['고객사명','제품','만료일','액션','일시'])
            self._load_logs()

        self.current_df   = None
        self.last_start   = None
        self.last_end     = None
        self.sort_column  = None
        self.sort_reverse = False
        self.rows         = []

        self.setup_widgets()
        self.load_filter_options()
        # 만료일 기준으로 이번달 데이터 로드
        self.load_this_month_preview()
        
        # 데이터 로딩 상태 출력
        print(f"RenewalMailerGUI 초기화 완료")
        print(f"raw_df 크기: {len(self.raw_df)}행")
        print(f"log_df 크기: {len(self.log_df)}행")
        print(f"current_df 크기: {len(self.current_df) if self.current_df is not None else 'None'}")
        print(f"rows 리스트 길이: {len(self.rows)}")
    
    def on_full_data_loaded(self):
        """전체 데이터 로딩 완료 시 호출되는 메서드"""
        if self.data_loader:
            try:
                # 데이터 로더에서 최신 데이터 가져오기
                self.raw_df = self.data_loader.get_renewal_data()
                self.log_df = self.data_loader.get_log_data()
                
                # 데이터가 비어있으면 기존 방식으로 로드
                if self.raw_df.empty:
                    print("Renewal 관리: 전체 데이터 로딩에서 데이터를 가져올 수 없어 기존 방식으로 로드합니다.")
                    self.raw_df = self._load_data_with_retry()
                    self.log_df = pd.DataFrame(columns=['고객사명','제품','만료일','액션','일시'])
                    self._load_logs()
                
                # 필터 옵션 다시 로드
                self.load_filter_options()
                # 현재 표시 중인 데이터 다시 로드
                if self.current_df is not None:
                    self.apply_filters()
                
                print(f"Renewal 관리: 전체 데이터 로딩 완료")
                print(f"raw_df 크기: {len(self.raw_df)}행")
            except Exception as e:
                print(f"Renewal 관리: 전체 데이터 로딩 오류 - {e}")
                # 에러 발생 시 기존 데이터 유지

    def _load_data_with_retry(self, max_retries=3):
        """데이터 로드 시 재시도 로직 - 최적화된 방식"""
        for attempt in range(max_retries):
            try:
                print(f"데이터 로드 시도 {attempt + 1}/{max_retries}")
                
                # 새로운 데이터 로더 사용
                if self.data_loader:
                    print("데이터 로더를 통한 로드 시도...")
                    data = self.data_loader.load_current_month_data()
                    if data and 'current_month' in data:
                        df = data['current_month']
                        # 컬럼명을 안전하게 문자열로 변환
                        df.columns = df.columns.astype(str).str.strip()
                        
                        # 날짜 컬럼 새로 생성
                        if '계산서 발행일' in df.columns:
                            df['계산서_날짜'] = pd.to_datetime(df['계산서 발행일'], errors='coerce')
                        if '만료일' in df.columns:
                            df['만료_날짜'] = pd.to_datetime(df['만료일'], errors='coerce')
                        
                        print(f"데이터 로더를 통한 로드 성공: {len(df)}행")
                        return df
                    else:
                        print("데이터 로더에서 데이터를 가져올 수 없습니다.")
                
                # 데이터 로더가 없거나 실패한 경우 기존 방식 사용
                print("기존 방식으로 데이터 로드 시도...")
                from core.data_loader import get_global_auth
                gc = get_global_auth()
                sh = gc.open_by_key('17rfmG5DEOj1CC-iA6SEVIdsS8zP-UC7lGBT6w2LwBvQ')
                try:
                    ws = sh.worksheet('renewal_list')
                    print("renewal_list 워크시트를 사용합니다.")
                except:
                    ws = sh.worksheet('Sheet1')
                    print("Sheet1 워크시트를 사용합니다.")
                df = pd.DataFrame(ws.get_all_records())
                
                # 컬럼명을 안전하게 문자열로 변환
                df.columns = df.columns.astype(str).str.strip()
                
                # 날짜 컬럼 새로 생성
                if '계산서 발행일' in df.columns:
                    df['계산서_날짜'] = pd.to_datetime(df['계산서 발행일'], errors='coerce')
                if '만료일' in df.columns:
                    df['만료_날짜'] = pd.to_datetime(df['만료일'], errors='coerce')
                
                print(f"기존 방식으로 데이터 로드 성공: {len(df)}행")
                return df
                
            except Exception as e:
                error_msg = str(e)
                print(f"데이터 로드 시도 {attempt + 1}/{max_retries} 실패: {error_msg}")
                import traceback
                traceback.print_exc()
                
                # 구글 API 관련 오류인 경우 더 긴 대기 시간
                if "500" in error_msg or "Internal error" in error_msg or "429" in error_msg:
                    wait_time = 5 * (attempt + 1)  # 5초, 10초, 15초
                    print(f"구글 API 오류 감지. {wait_time}초 대기 후 재시도...")
                else:
                    wait_time = 2
                    
                if attempt < max_retries - 1:
                    import time
                    time.sleep(wait_time)
                else:
                    print("최대 재시도 횟수 초과. 빈 DataFrame으로 초기화합니다.")
                    return pd.DataFrame()
        
        return pd.DataFrame()

    def _load_logs(self):
        """로그 시트를 불러와 self.log_df에 저장 - 최적화된 방식"""
        try:
            # 새로운 데이터 로더 사용
            if self.data_loader:
                data = self.data_loader.load_current_month_data()
                if data and 'log_data' in data:
                    self.log_df = data['log_data']
                    if not self.log_df.empty:
                        self.log_df.columns = self.log_df.columns.astype(str).str.strip()
                        if '일시' in self.log_df.columns:
                            self.log_df['일시'] = pd.to_datetime(
                                self.log_df['일시'], format='%Y-%m-%d %H:%M', errors='coerce'
                            )
                            self.log_df.sort_values('일시', ascending=False, inplace=True)
                    return
            
            # 데이터 로더가 없거나 실패한 경우 기존 방식 사용
            from core.data_loader import get_global_auth
            gc = get_global_auth()
            ws = gc.open_by_key(LOG_SHEET_ID).worksheet('Sheet1')
            logs = ws.get_all_records()
            self.log_df = pd.DataFrame(logs)
            if not self.log_df.empty:
                self.log_df.columns = self.log_df.columns.astype(str).str.strip()
                if '일시' in self.log_df.columns:
                    self.log_df['일시'] = pd.to_datetime(
                        self.log_df['일시'], format='%Y-%m-%d %H:%M', errors='coerce'
                    )
                    self.log_df.sort_values('일시', ascending=False, inplace=True)
        except Exception as e:
            print(f"로그 데이터 로드 오류: {e}")
            self.log_df = pd.DataFrame(columns=['고객사명','제품','만료일','액션','일시'])

    def setup_widgets(self):
        # 상단 헤더
        header_frame = tk.Frame(self.root, bg="#F7F9FB", height=60)
        header_frame.pack(fill='x', pady=(0, 20))
        header_frame.pack_propagate(False)
        
        # 헤더 제목
        title_label = tk.Label(header_frame, text="Renewal 관리", 
                              font=("맑은 고딕", 16, "bold"), 
                              fg="#1F2937", bg="#F7F9FB")
        title_label.pack(side='left', padx=20, pady=15)
        
        # 사용자 정보 (우측)
        user_label = tk.Label(header_frame, text=f"👤 {self.username}", 
                             font=("맑은 고딕", 11), 
                             fg="#6B7280", bg="#F7F9FB")
        user_label.pack(side='right', padx=20, pady=15)

        # 검색/필터 카드 섹션
        filter_section = tk.Frame(self.root, bg="#F7F9FB")
        filter_section.pack(fill='x', padx=20, pady=(0, 20))

        # 첫 번째 카드: 기준 선택 및 날짜
        card1 = tk.Frame(filter_section, bg="white", relief="solid", bd=1)
        card1.pack(side='left', fill='both', expand=True, padx=(0, 10))
        # 카드1 헤더
        card1_header = tk.Frame(card1, bg="#F8FAFC", height=40)
        card1_header.pack(fill='x')
        card1_header.pack_propagate(False)
        tk.Label(card1_header, text="📅 조회 기준", font=("맑은 고딕", 11, "bold"), 
                fg="#374151", bg="#F8FAFC").pack(side='left', padx=15, pady=10)
        # 카드1 내용
        card1_content = tk.Frame(card1, bg="white")
        card1_content.pack(fill='both', expand=True, padx=15, pady=15)
        
        # 기준 선택
        criteria_frame = tk.Frame(card1_content, bg="white")
        criteria_frame.pack(fill='x', pady=(0, 10))
        tk.Label(criteria_frame, text="조회 기준:", font=("맑은 고딕", 10), 
                fg="#374151", bg="white").pack(side='left')
        tk.Radiobutton(criteria_frame, text="계산서 발행일", variable=self.criteria_var, 
                      value="계산서 발행일", command=self.on_criteria_changed,
                      font=("맑은 고딕", 10), fg="#374151", bg="white").pack(side='left', padx=5)
        tk.Radiobutton(criteria_frame, text="만료일", variable=self.criteria_var, 
                      value="만료일", command=self.on_criteria_changed,
                      font=("맑은 고딕", 10), fg="#374151", bg="white").pack(side='left', padx=5)
        
        # 날짜 선택
        date_frame = tk.Frame(card1_content, bg="white")
        date_frame.pack(fill='x', pady=(0, 10))
        tk.Label(date_frame, text="조회 기준일:", font=("맑은 고딕", 10), 
                fg="#374151", bg="white").pack(side='left')
        self.date_entry = DateEntry(date_frame, width=12, date_pattern='yyyy-mm-dd')
        self.date_entry.set_date(datetime.today())
        self.date_entry.pack(side='left', padx=5)
        self.date_entry.bind("<<DateEntrySelected>>", lambda e: self.load_this_month_preview())

        # 기간 버튼들 - 2줄로 분리
        # 첫 번째 줄: 이번달, 다음달, 미래 1년
        first_row_frame = tk.Frame(card1_content, bg="white")
        first_row_frame.pack(fill='x', pady=(0, 5))
        
        first_row_periods = [
            ("이번달",   self.load_this_month_preview),
            ("다음달",   self.load_next_month_preview),
            ("미래 1년", self.load_one_year_preview),
        ]
        for text, cmd in first_row_periods:
            btn = tk.Button(
                first_row_frame, text=text,
                command=cmd,
                font=("맑은 고딕", 9), bg="#3B82F6", fg="white",
                relief="flat", padx=10, pady=3
            )
            btn.pack(side='left', padx=2)
            # hover 효과
            btn.bind('<Enter>', lambda e, b=btn: b.configure(bg="#2563EB"))
            btn.bind('<Leave>', lambda e, b=btn: b.configure(bg="#3B82F6"))
        
        # 두 번째 줄: 과거 1년, 새로고침
        second_row_frame = tk.Frame(card1_content, bg="white")
        second_row_frame.pack(fill='x')
        
        # 과거 1년 버튼
        past_year_btn = tk.Button(second_row_frame, text="과거 1년", 
                                 command=self.load_past_year_preview,
                                 font=("맑은 고딕", 9), bg="#3B82F6", fg="white",
                                 relief="flat", padx=10, pady=3)
        past_year_btn.pack(side='left', padx=2)
        past_year_btn.bind('<Enter>', lambda e, b=past_year_btn: b.configure(bg="#2563EB"))
        past_year_btn.bind('<Leave>', lambda e, b=past_year_btn: b.configure(bg="#3B82F6"))
        
        # 새로고침 버튼
        refresh_btn = tk.Button(second_row_frame, text="🔄 새로고침",
                               command=self.refresh_data,
                               font=("맑은 고딕", 9), bg="#10B981", fg="white",
                               relief="flat", padx=10, pady=3)
        refresh_btn.pack(side='left', padx=2)
        refresh_btn.bind('<Enter>', lambda e, b=refresh_btn: b.configure(bg="#059669"))
        refresh_btn.bind('<Leave>', lambda e, b=refresh_btn: b.configure(bg="#10B981"))
        
        # 전체 데이터 로드 버튼 (A)
        all_data_btn = tk.Button(second_row_frame, text="A", command=self.load_all_data,
                                font=("맑은 고딕", 9, "bold"), bg="#8B5CF6", fg="white",
                                relief="flat", padx=10, pady=3)
        all_data_btn.pack(side='left', padx=2)
        # hover 효과
        all_data_btn.bind('<Enter>', lambda e, b=all_data_btn: b.configure(bg="#7C3AED"))
        all_data_btn.bind('<Leave>', lambda e, b=all_data_btn: b.configure(bg="#8B5CF6"))

        # 두 번째 카드: 필터 옵션
        card2 = tk.Frame(filter_section, bg="white", relief="solid", bd=1)
        card2.pack(side='left', fill='both', expand=True, padx=(0, 10))
        # 카드2 헤더
        card2_header = tk.Frame(card2, bg="#F8FAFC", height=40)
        card2_header.pack(fill='x')
        card2_header.pack_propagate(False)
        tk.Label(card2_header, text="🔍 필터", font=("맑은 고딕", 11, "bold"), 
                fg="#374151", bg="#F8FAFC").pack(side='left', padx=15, pady=10)
        # 카드2 내용
        card2_content = tk.Frame(card2, bg="white")
        card2_content.pack(fill='both', expand=True, padx=15, pady=15)
        
        # 첫 번째 행: 고객사명, 거래처명
        row1 = tk.Frame(card2_content, bg="white")
        row1.pack(fill='x', pady=(0, 10))
        
        tk.Label(row1, text="고객사명:", font=("맑은 고딕", 10), 
                fg="#374151", bg="white").pack(side='left')
        # 안전하게 고객사명 리스트 생성
        if not self.raw_df.empty and '고객사명' in self.raw_df.columns:
            try:
                self.all_customers = sorted(self.raw_df['고객사명'].dropna().unique().tolist())
            except Exception as e:
                print(f"고객사명 리스트 생성 오류: {e}")
                self.all_customers = []
        else:
            self.all_customers = []
        self.customer_cb = ttk.Combobox(
            row1,
            state='normal',
            width=20,
            postcommand=self.update_customer_list
        )
        self.customer_cb.pack(side='left', padx=5)
        self.customer_cb['values'] = ['전체'] + self.all_customers
        self.customer_cb.set('전체')
        self.customer_cb.bind('<<ComboboxSelected>>', lambda e: self.apply_filters())
        self.customer_cb.bind('<Return>', lambda e: self.on_customer_search())

        tk.Label(row1, text="거래처명:", font=("맑은 고딕", 10), 
                fg="#374151", bg="white").pack(side='left', padx=(20,0))
        # 안전하게 거래처명 리스트 생성
        if not self.raw_df.empty and '거래처명' in self.raw_df.columns:
            try:
                self.all_clients = sorted(self.raw_df['거래처명'].dropna().unique().tolist())
            except Exception as e:
                print(f"거래처명 리스트 생성 오류: {e}")
                self.all_clients = []
        else:
            self.all_clients = []
        self.client_cb = ttk.Combobox(
            row1,
            state='normal',
            width=20,
            postcommand=self.update_client_list
        )
        self.client_cb.pack(side='left', padx=5)
        self.client_cb['values'] = ['전체'] + self.all_clients
        self.client_cb.set('전체')
        self.client_cb.bind('<<ComboboxSelected>>', lambda e: self.apply_filters())
        self.client_cb.bind('<Return>', lambda e: self.on_client_search())

        # 두 번째 행: 대분류, 영업사원
        row2 = tk.Frame(card2_content, bg="white")
        row2.pack(fill='x', pady=(0, 10))
        
        tk.Label(row2, text="대분류:", font=("맑은 고딕", 10), 
                fg="#374151", bg="white").pack(side='left')
        self.category_cb = ttk.Combobox(row2, state='readonly', width=15)
        self.category_cb.pack(side='left', padx=5)
        self.category_cb.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        tk.Label(row2, text="영업사원:", font=("맑은 고딕", 10), 
                fg="#374151", bg="white").pack(side='left', padx=(20,0))
        self.sales_cb = ttk.Combobox(row2, state='readonly', width=15)
        self.sales_cb.pack(side='left', padx=5)
        self.sales_cb.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        # 세 번째 행: 검색 버튼만
        row3 = tk.Frame(card2_content, bg="white")
        row3.pack(fill='x')
        
        search_btn = tk.Button(row3, text="🔍 검색", command=self.apply_filters,
                              font=("맑은 고딕", 10, "bold"), bg="#10B981", fg="white",
                              relief="flat", padx=15, pady=4)
        search_btn.pack(side='left')
        search_btn.bind('<Enter>', lambda e: search_btn.configure(bg="#059669"))
        search_btn.bind('<Leave>', lambda e: search_btn.configure(bg="#10B981"))

        # 세 번째 카드: 액션 버튼
        card3 = tk.Frame(filter_section, bg="white", relief="solid", bd=1)
        card3.pack(side='left', fill='both', expand=True)
        # 카드3 헤더 및 내용 기존대로 card3에 배치
        # 카드3 헤더
        card3_header = tk.Frame(card3, bg="#F8FAFC", height=40)
        card3_header.pack(fill='x')
        card3_header.pack_propagate(False)
        tk.Label(card3_header, text="⚡ 액션", font=("맑은 고딕", 11, "bold"), 
                fg="#374151", bg="#F8FAFC").pack(side='left', padx=15, pady=10)
        
        # 카드3 내용
        card3_content = tk.Frame(card3, bg="white")
        card3_content.pack(fill='both', expand=True, padx=15, pady=15)
        
        # 액션 버튼들 (세로 배치)
        action_buttons = tk.Frame(card3_content, bg="white")
        action_buttons.pack(fill='both', expand=True)
        
        # 메일 작성 버튼
        mail_btn = tk.Button(action_buttons, text="📧 메일 작성", command=self.send_selected_email,
                            font=("맑은 고딕", 10, "bold"), bg="#3B82F6", fg="white",
                            relief="flat", padx=15, pady=4)
        mail_btn.pack(fill='x', pady=(0, 5))
        mail_btn.bind('<Enter>', lambda e: mail_btn.configure(bg="#2563EB"))
        mail_btn.bind('<Leave>', lambda e: mail_btn.configure(bg="#3B82F6"))
        
        # 견적서 첨부 버튼
        quote_btn = tk.Button(action_buttons, text="📋 견적서 첨부", command=self.quote_selected,
                             font=("맑은 고딕", 10, "bold"), bg="#8B5CF6", fg="white",
                             relief="flat", padx=15, pady=4)
        quote_btn.pack(fill='x', pady=(0, 5))
        quote_btn.bind('<Enter>', lambda e: quote_btn.configure(bg="#7C3AED"))
        quote_btn.bind('<Leave>', lambda e: quote_btn.configure(bg="#8B5CF6"))
        
        # 로그 보기 버튼
        log_btn = tk.Button(action_buttons, text="📊 로그 보기", command=lambda: show_logs(self.root),
                           font=("맑은 고딕", 10, "bold"), bg="#6B7280", fg="white",
                           relief="flat", padx=15, pady=4)
        log_btn.pack(fill='x')
        log_btn.bind('<Enter>', lambda e: log_btn.configure(bg="#4B5563"))
        log_btn.bind('<Leave>', lambda e: log_btn.configure(bg="#6B7280"))

        # 데이터 테이블 섹션
        table_section = tk.Frame(self.root, bg="#F7F9FB")
        table_section.pack(fill='both', expand=True, padx=20, pady=(0, 20))
        
        # 테이블 헤더
        table_header = tk.Frame(table_section, bg="white", relief="solid", bd=1)
        table_header.pack(fill='x', pady=(0, 1))
        tk.Label(table_header, text="📊 데이터 목록", font=("맑은 고딕", 14, "bold"), 
                fg="#1F2937", bg="white").pack(side='left', padx=15, pady=10)

        # 테이블 컨테이너
        self.cols = (
            "액션","계산서 발행일","고객사명","거래처명","영업담당자",
            "대분류","제품","수량","만료일",
            "판매 단가","판매 합계","원가","원가계","이익액","이익율"
        )
        container = tk.Frame(table_section, bg="white", relief="solid", bd=1)
        container.pack(fill='both', expand=True)
        vsb = ttk.Scrollbar(container, orient='vertical'); vsb.pack(side='right', fill='y')
        hsb = ttk.Scrollbar(container, orient='horizontal'); hsb.pack(side='bottom', fill='x')

        self.tree = ttk.Treeview(
            container, columns=self.cols, show='headings', selectmode='browse',
            yscrollcommand=vsb.set, xscrollcommand=hsb.set
        )
        vsb.config(command=self.tree.yview); hsb.config(command=self.tree.xview)

        for c in self.cols:
            self.tree.heading(c, text=c, command=lambda col=c: self.sort_by_column(col))
            if c in ("고객사명","영업담당자","대분류","제품"):
                anchor = 'w'
            elif c in ("판매 단가","판매 합계","원가","원가계","이익액"):
                anchor = 'e'
            else:
                anchor = 'center'
            self.tree.column(c, width=120, anchor=anchor, stretch=False)

        self.tree.pack(fill='both', expand=True, padx=1, pady=1)

        # 컬럼 크기 조절 기능 추가
        from ui.resizable_treeview import ResizableTreeview
        self.resizable_tree = ResizableTreeview(self.tree, "renewal_table")

        # Shift+휠로 가로 스크롤 빠르게 이동
        def fast_horizontal_scroll(event):
            try:
                widget = event.widget
                if widget and widget.winfo_exists():
                    if event.delta < 0:
                        widget.xview_scroll(3, "units")
                    else:
                        widget.xview_scroll(-3, "units")
            except Exception:
                # 위젯이 파괴되었거나 접근할 수 없는 경우 무시
                pass
        self.tree.bind("<Shift-MouseWheel>", fast_horizontal_scroll)
        self.tree.bind("<Double-1>", self.show_detail)

    def on_criteria_changed(self):
        """기준이 변경되었을 때 현재 필터를 다시 적용"""
        print(f"조회 기준 변경됨: '{self.criteria_var.get()}'")
        if self.last_start is not None and self.last_end is not None:
            print(f"마지막 필터링 기간으로 재적용: {self.last_start} ~ {self.last_end}")
            self._filter_cached_data(self.last_start, self.last_end)
        else:
            print("마지막 필터링 기간이 없어 재적용하지 않습니다.")

    def load_filter_options(self):
        try:
            df_all = self.raw_df.copy()
            # 컬럼명을 안전하게 문자열로 변환
            if not df_all.empty:
                df_all.columns = df_all.columns.astype(str).str.strip()
                
                if '대분류' in df_all.columns:
                    cats = sorted(df_all['대분류'].dropna().unique().tolist())
                else:
                    cats = []
                    
                if '영업사원' in df_all.columns:
                    sales = sorted(df_all['영업사원'].dropna().unique().tolist())
                else:
                    sales = []
                    
                if '고객사명' in df_all.columns:
                    self.all_customers = sorted(df_all['고객사명'].dropna().unique().tolist())
                else:
                    self.all_customers = []
                    
                if '거래처명' in df_all.columns:
                    self.all_clients = sorted(df_all['거래처명'].dropna().unique().tolist())
                else:
                    self.all_clients = []
            else:
                cats = []
                sales = []
                self.all_customers = []
                self.all_clients = []
                
            self.category_cb['values'] = ['전체'] + cats
            self.category_cb.current(0)
            
            # 고객사명과 거래처명 콤보박스 값 업데이트
            self.customer_cb['values'] = ['전체'] + self.all_customers
            self.client_cb['values'] = ['전체'] + self.all_clients
            
        except Exception as e:
            print("필터 옵션 생성 오류:", e)
            self.category_cb['values'] = ['전체']
            self.category_cb.current(0)
            sales = []
            self.all_customers = []
            self.all_clients = []
            
        self.sales_cb['values'] = ['전체'] + sales
        self.sales_cb.current(0)

    def _filter_cached_data(self, start, end):
        self.last_start, self.last_end = start, end
        df = self.raw_df.copy()
        
        print(f"기간 필터링 시작: 원본 데이터 {len(df)}행")
        print(f"필터링 기간: {start} ~ {end}")
        print(f"조회 기준: {self.criteria_var.get()}")
        
        # 빈 DataFrame인 경우 처리
        if df.empty:
            print("경고: 원본 데이터가 비어있습니다.")
            self.current_df = df
            self.populate_tree(df)
            return
            
        # 컬럼명을 안전하게 문자열로 변환 후 strip 적용
        df.columns = df.columns.astype(str).str.strip()
        
        # 데이터 로더에서 가져온 데이터인지 확인 (이미 날짜 컬럼이 변환되어 있을 수 있음)
        has_converted_dates = '계산서_날짜' in df.columns or '만료_날짜' in df.columns
        
        if not has_converted_dates:
            # 기존 방식: 날짜 컬럼 변환
            if '만료일' in df.columns:
                df['만료일'] = try_parse_date(df['만료일'].astype(str).str.strip())
            if '계산서 발행일' in df.columns:
                df['계산서 발행일'] = try_parse_date(df['계산서 발행일'].astype(str).str.strip())
                
            # 날짜 변환: 새 컬럼에만 적용 (컬럼 존재 여부 확인)
            if '계산서 발행일' in df.columns:
                df['계산서_날짜'] = try_parse_date(df['계산서 발행일'].astype(str).str.strip())
            else:
                df['계산서_날짜'] = pd.NaT
            if '만료일' in df.columns:
                df['만료_날짜'] = try_parse_date(df['만료일'].astype(str).str.strip())
            else:
                df['만료_날짜'] = pd.NaT
        else:
            print("데이터 로더에서 이미 변환된 날짜 컬럼을 사용합니다.")
            
        start_dt = pd.to_datetime(start)
        end_dt = pd.to_datetime(end)
        
        print(f"변환된 날짜 범위: {start_dt} ~ {end_dt}")
        
        # 날짜 필터링 - 조회 기준에 따라 다르게 처리
        criteria = self.criteria_var.get()
        print(f"선택된 조회 기준: '{criteria}'")
        
        if criteria == "계산서 발행일":
            if '계산서_날짜' in df.columns:
                # NaT 값 제외하고 필터링
                valid_mask = df['계산서_날짜'].notna()
                date_mask = (df['계산서_날짜'] >= start_dt) & (df['계산서_날짜'] <= end_dt)
                mask = valid_mask & date_mask
                sel_df = df.loc[mask]
                print(f"계산서 발행일 기준 필터링: {len(df)}행 → {len(sel_df)}행")
                print(f"유효한 계산서 발행일 데이터: {valid_mask.sum()}행")
                print(f"기간 내 계산서 발행일 데이터: {date_mask.sum()}행")
            else:
                print("경고: 계산서 발행일 컬럼이 없습니다.")
                sel_df = df
        elif criteria == "만료일":
            if '만료_날짜' in df.columns:
                # NaT 값 제외하고 필터링
                valid_mask = df['만료_날짜'].notna()
                date_mask = (df['만료_날짜'] >= start_dt) & (df['만료_날짜'] <= end_dt)
                mask = valid_mask & date_mask
                sel_df = df.loc[mask]
                print(f"만료일 기준 필터링: {len(df)}행 → {len(sel_df)}행")
                print(f"유효한 만료일 데이터: {valid_mask.sum()}행")
                print(f"기간 내 만료일 데이터: {date_mask.sum()}행")
            else:
                print("경고: 만료일 컬럼이 없습니다.")
                sel_df = df
        else:
            print(f"알 수 없는 조회 기준: '{criteria}', 전체 데이터 사용")
            sel_df = df
            
        # 필터링 결과 확인
        if len(sel_df) == 0:
            print("경고: 해당 기간에 데이터가 없습니다.")
            if '계산서_날짜' in df.columns:
                valid_dates = df['계산서_날짜'].dropna()
                if len(valid_dates) > 0:
                    print(f"계산서 발행일 범위: {valid_dates.min()} ~ {valid_dates.max()}")
                else:
                    print("계산서 발행일 데이터가 없습니다.")
            if '만료_날짜' in df.columns:
                valid_dates = df['만료_날짜'].dropna()
                if len(valid_dates) > 0:
                    print(f"만료일 범위: {valid_dates.min()} ~ {valid_dates.max()}")
                else:
                    print("만료일 데이터가 없습니다.")
        
        self.current_df = sel_df
        self.populate_tree(self.current_df)

    def load_this_month_preview(self):
        print("이번달 버튼 클릭")
        self.clear_search_fields()
        base = self.date_entry.get_date()
        start = base.replace(day=1)
        import calendar
        last_day = calendar.monthrange(base.year, base.month)[1]
        end = base.replace(day=last_day)
        print(f"이번달 기간: {start} ~ {end}")
        self._filter_cached_data(pd.to_datetime(start), pd.to_datetime(end))

    def load_next_month_preview(self):
        print("다음달 버튼 클릭")
        self.clear_search_fields()
        base = self.date_entry.get_date()
        year = base.year + (base.month // 12)
        month = (base.month % 12) + 1
        start = base.replace(year=year, month=month, day=1)
        import calendar
        last_day = calendar.monthrange(year, month)[1]
        end = start.replace(day=last_day)
        print(f"다음달 기간: {start} ~ {end}")
        self._filter_cached_data(pd.to_datetime(start), pd.to_datetime(end))

    def load_one_year_preview(self):
        print("미래 1년 버튼 클릭")
        self.clear_search_fields()
        base = self.date_entry.get_date()
        start = pd.to_datetime(base)
        end = start + pd.DateOffset(years=1) - pd.Timedelta(days=1)
        print(f"미래 1년 기간: {start} ~ {end}")
        self._filter_cached_data(start, end)

    def load_past_year_preview(self):
        print("과거 1년 버튼 클릭")
        self.clear_search_fields()
        base = self.date_entry.get_date()
        end = pd.to_datetime(base)
        start = end - pd.DateOffset(years=1) + pd.Timedelta(days=1)
        print(f"과거 1년 기간: {start} ~ {end}")
        self._filter_cached_data(start, end)

    def refresh_data(self):
        """데이터 새로고침"""
        print("새로고침 버튼 클릭")
        try:
            # 현재 필터 상태를 유지하면서 데이터 다시 로드
            if hasattr(self, 'last_start') and hasattr(self, 'last_end') and self.last_start is not None and self.last_end is not None:
                print(f"마지막 필터링 기간으로 데이터 새로고침: {self.last_start} ~ {self.last_end}")
                self._filter_cached_data(self.last_start, self.last_end)
            else:
                print("마지막 필터링 기간이 없어 이번달 데이터로 새로고침")
                self.load_this_month_preview()
        except Exception as e:
            print(f"데이터 새로고침 오류: {e}")
            messagebox.showerror("오류", f"데이터 새로고침 중 오류가 발생했습니다:\n{e}")

    def load_all_data(self):
        """전체 데이터 로드 (기간 제한 없음)"""
        print("전체 데이터 로드 버튼 클릭")
        try:
            # 필터 필드 초기화
            self.clear_search_fields()
            
            # 전체 데이터를 트리뷰에 표시
            if self.raw_df is not None and not self.raw_df.empty:
                print(f"전체 데이터 로드: {len(self.raw_df)}행")
                # current_df 설정 (정렬을 위해 필요)
                self.current_df = self.raw_df.copy()
                self.populate_tree(self.raw_df)
                messagebox.showinfo("완료", f"전체 데이터 {len(self.raw_df)}행을 로드했습니다.")
            else:
                print("로드할 데이터가 없습니다.")
                messagebox.showwarning("경고", "로드할 데이터가 없습니다.")
        except Exception as e:
            print(f"전체 데이터 로드 오류: {e}")
            messagebox.showerror("오류", f"전체 데이터 로드 중 오류가 발생했습니다:\n{e}")

    def clear_search_fields(self):
        self.category_cb.current(0)
        self.sales_cb.current(0)
        self.customer_cb.set('전체')
        self.client_cb.set('전체')

    def apply_filters(self):
        if self.current_df is None:
            return
        df = self.current_df.copy() if self.current_df is not None else pd.DataFrame()
        
        if df.empty:
            self.populate_tree(df)
            return
            
        print(f"필터 적용 시작: 원본 데이터 {len(df)}행")
        
        try:
            # 대분류 필터링
            if self.category_cb.get() != '전체' and '대분류' in df.columns:
                before_count = len(df)
                df = df[df['대분류'] == self.category_cb.get()]
                print(f"대분류 '{self.category_cb.get()}' 필터링: {before_count}행 → {len(df)}행")
                
            # 영업사원 필터링
            if self.sales_cb.get() != '전체' and '영업사원' in df.columns:
                before_count = len(df)
                df = df[df['영업사원'] == self.sales_cb.get()]
                print(f"영업사원 '{self.sales_cb.get()}' 필터링: {before_count}행 → {len(df)}행")
                
            # 고객사명 필터링
            cust = self.customer_cb.get().strip()
            if cust and cust != '전체' and '고객사명' in df.columns:
                before_count = len(df)
                # 대소문자 구분 없이 검색하고, NaN 값 처리
                df = df[df['고객사명'].str.contains(cust, na=False, case=False, regex=False)]
                print(f"고객사명 '{cust}' 필터링: {before_count}행 → {len(df)}행")
                if len(df) == 0:
                    print(f"경고: 고객사명 '{cust}'에 해당하는 데이터가 없습니다.")
                    print(f"사용 가능한 고객사명 샘플: {df['고객사명'].dropna().unique()[:5].tolist()}")
                    
            # 거래처명 필터링
            client = self.client_cb.get().strip()
            if client and client != '전체' and '거래처명' in df.columns:
                before_count = len(df)
                # 대소문자 구분 없이 검색하고, NaN 값 처리
                df = df[df['거래처명'].str.contains(client, na=False, case=False, regex=False)]
                print(f"거래처명 '{client}' 필터링: {before_count}행 → {len(df)}행")
                if len(df) == 0:
                    print(f"경고: 거래처명 '{client}'에 해당하는 데이터가 없습니다.")
                    
        except Exception as e:
            print(f"필터 적용 오류: {e}")
            import traceback
            traceback.print_exc()
            
        print(f"필터 적용 완료: 최종 {len(df)}행")
        self.populate_tree(df)

    def populate_tree(self, df):
        self.tree.delete(*self.tree.get_children())
        self.rows = []
        
        for idx, row in df.iterrows():
            # 만료일, 계산서 발행일 날짜 포맷팅 (공백/NaT/None은 빈칸)
            expiry = ''
            if '만료일' in row and pd.notna(row['만료일']):
                if isinstance(row['만료일'], pd.Timestamp):
                    expiry = row['만료일'].strftime('%Y-%m-%d')
                elif isinstance(row['만료일'], str) and row['만료일'].strip():
                    # 시간 부분이 포함된 경우 날짜만 추출
                    expiry = row['만료일'].split(' ')[0] if ' ' in row['만료일'] else row['만료일']
                elif isinstance(row['만료일'], (datetime.date, datetime.datetime)):
                    expiry = row['만료일'].strftime('%Y-%m-%d')
            
            bill_date = ''
            if '계산서 발행일' in row and pd.notna(row['계산서 발행일']):
                if isinstance(row['계산서 발행일'], pd.Timestamp):
                    bill_date = row['계산서 발행일'].strftime('%Y-%m-%d')
                elif isinstance(row['계산서 발행일'], str) and row['계산서 발행일'].strip():
                    # 시간 부분이 포함된 경우 날짜만 추출
                    bill_date = row['계산서 발행일'].split(' ')[0] if ' ' in row['계산서 발행일'] else row['계산서 발행일']
                elif isinstance(row['계산서 발행일'], (datetime.date, datetime.datetime)):
                    bill_date = row['계산서 발행일'].strftime('%Y-%m-%d')
            # 이하 기존 코드에서 row['만료일'], row['계산서 발행일'] 대신 expiry, bill_date 사용
            if not self.log_df.empty:
                # 로그 데이터의 날짜 형식 통일
                log_expiry = ''
                if '만료일' in self.log_df.columns:
                    # 로그의 만료일을 문자열로 변환하여 비교
                    log_expiry_col = self.log_df['만료일'].astype(str)
                    # 날짜 형식 통일 (YYYY-MM-DD)
                    log_expiry_col = log_expiry_col.apply(lambda x: x.split(' ')[0] if ' ' in str(x) else str(x))
                
                # 매칭 조건: 고객사명, 제품, 만료일
                mask = (
                    (self.log_df['고객사명'] == row['고객사명']) &
                    (self.log_df['제품'] == row['제품'])
                )
                
                # 만료일 매칭 (더 유연하게)
                if '만료일' in self.log_df.columns and expiry:
                    # 로그의 만료일과 현재 행의 만료일 비교
                    date_mask = log_expiry_col == expiry
                    mask = mask & date_mask
                
                acts = self.log_df.loc[mask, '액션']
                mark = acts.iloc[-1] if len(acts) > 0 else ''
                uc = row.get('실 매입 단가')  if pd.notna(row.get('실 매입 단가'))  else row.get('원가',0)
                sc = row.get('실 매입 합계') if pd.notna(row.get('실 매입 합계')) else row.get('원가계',0)
                pv = row.get('실 이익액')    if pd.notna(row.get('실 이익액'))    else row.get('이익액',0)
                pr = row.get('실 이익율')    if pd.notna(row.get('실 이익율'))    else row.get('이익율',0)
                unit_cost_str = f"{safe_int_convert(uc):,}"
                sum_cost_str  = f"{safe_int_convert(sc):,}"
                profit_str    = f"{safe_int_convert(pv):,}"
                rate_str      = pr if isinstance(pr, str) else f"{pr:.2%}"
                vals = (
                    mark, bill_date,
                    row['고객사명'], row.get('거래처명',''), row.get('영업사원',''),
                    row.get('대분류',''), row['제품'], row['수량'], expiry,
                    f"{safe_int_convert(row.get('판매 단가',0)):,}", f"{safe_int_convert(row.get('판매 합계',0)):,}",
                    unit_cost_str, sum_cost_str, profit_str, rate_str
                )
                self.tree.insert('', 'end', iid=str(idx), values=vals)
                # 원본 row를 복사하되, 날짜 필드를 안전하게 처리
                safe_row = row.copy()
                if '만료일' in safe_row and pd.isna(safe_row['만료일']):
                    safe_row['만료일'] = None
                if '계산서 발행일' in safe_row and pd.isna(safe_row['계산서 발행일']):
                    safe_row['계산서 발행일'] = None
                self.rows.append(safe_row)
        

    def sort_by_column(self, col):
        if col == '액션': return
        df = self.current_df.copy()
        rev = (col == self.sort_column) and not self.sort_reverse
        self.sort_column, self.sort_reverse = col, rev
        
        print(f"정렬 시작: 컬럼={col}, 역순={rev}, 데이터 행 수={len(df)}")
        
        # 날짜 컬럼인 경우 날짜 형식으로 변환 후 정렬
        if col in ("만료일", "계산서 발행일"):
            # 날짜 컬럼을 datetime으로 변환
            df[col] = pd.to_datetime(df[col], errors='coerce')
            # NaT 값을 마지막으로 보내기 위해 fillna 사용
            df = df.sort_values(col, ascending=not rev, na_position='last')
            print(f"날짜 정렬 완료: {col}, 첫 번째 값={df[col].iloc[0] if not df.empty else 'None'}")
        elif col in ("수량", "판매 단가", "판매 합계", "원가", "원가계", "이익액", "이익율"):
            # 숫자 컬럼인 경우 숫자로 변환 후 정렬
            df[col] = pd.to_numeric(df[col], errors='coerce')
            df = df.sort_values(col, ascending=not rev, na_position='last')
            print(f"숫자 정렬 완료: {col}")
        else:
            # 문자열 컬럼인 경우 문자열로 정렬
            df = df.sort_values(col, ascending=not rev, key=lambda x: x.fillna('').astype(str))
            print(f"문자열 정렬 완료: {col}")
        
        print(f"정렬 완료: 최종 데이터 행 수={len(df)}")
        self.populate_tree(df)

    def show_detail(self, event):
        sel = self.tree.selection()
        if not sel: return
        pos = self.tree.index(sel[0])
        if 0 <= pos < len(self.rows):
            show_customer_detail(self.root, self.rows[pos])

    def send_selected_email(self):
        sel = self.tree.selection()
        if len(sel) != 1:
            messagebox.showwarning("경고", "메일 작성은 한 행만 선택해야 합니다.")
            return
        
        iid = sel[0]
        pos = self.tree.index(iid)
        
        if pos >= len(self.rows):
            messagebox.showerror("오류", f"데이터 인덱스 오류: {pos} >= {len(self.rows)}")
            return
            
        row = self.rows[pos]
        
        # 아웃룩 연결 상태 미리 확인
        try:
            from email_module.renewal_email import check_outlook_connection
            if not check_outlook_connection():
                messagebox.showerror("아웃룩 연결 오류", 
                    "아웃룩에 연결할 수 없습니다.\n\n아웃룩을 실행한 후 다시 시도해주세요.")
                return
        except Exception as e:
            # 연결 확인에 실패해도 계속 진행
            pass
        
        # 만료일 안전 처리
        expiry_date = safe_parse_date(row['만료일'])
        if expiry_date is None:
            messagebox.showerror("오류", "만료일이 유효하지 않습니다. 데이터를 확인해주세요.")
            return
        
        sel_month = expiry_date.month
        client = row.get('거래처명','')
        same = [j for j,r in enumerate(self.rows)
                if safe_parse_date(r['만료일']) is not None and 
                safe_parse_date(r['만료일']).month == sel_month and 
                r.get('거래처명','')==client]
        if len(same) >= 2:
            dlg = tk.Toplevel(self.root); dlg.title("유통건 확인")
            tk.Label(dlg, text="2건 이상이 같은달에 있는 건입니다. 어떻게 하시겠어요?", padx=20, pady=10).pack()
            def single():
                dlg.destroy()
                res_email.send_emails([row], self.username)
                self.tree.set(iid, '액션', '메일 작성')
                messagebox.showinfo("완료","메일 작성 로그가 기록되었습니다.")
            def multi():
                dlg.destroy()
                lst = [self.rows[j] for j in same]
                send_multi_emails(lst, self.username)
                for j in same:
                    child_iid = self.tree.get_children()[j]
                    self.tree.set(child_iid, '액션', '메일 작성')
                messagebox.showinfo("완료", f"{len(same)}건 메일 작성 로그가 기록되었습니다.")
            frm = tk.Frame(dlg); frm.pack(pady=10)
            tk.Button(frm, text="1건만 발송", width=12, command=single).pack(side='left', padx=5)
            tk.Button(frm, text="2건 이상 발송", width=12, command=multi).pack(side='right', padx=5)
            dlg.transient(self.root); dlg.grab_set(); self.root.wait_window(dlg)
            return
        res_email.send_emails([row], self.username)
        self.tree.set(iid, '액션', '메일 작성')
        messagebox.showinfo("완료","메일 작성 로그가 기록되었습니다.")

    def quote_selected(self):
        sel = self.tree.selection()
        if len(sel) != 1:
            messagebox.showwarning("경고", "견적서 첨부는 한 행만 선택해야 합니다.")
            return
        
        iid = sel[0]
        pos = self.tree.index(iid)
        
        if pos >= len(self.rows):
            messagebox.showerror("오류", f"견적서 데이터 인덱스 오류: {pos} >= {len(self.rows)}")
            return
            
        row = self.rows[pos]
        
        # 아웃룩 연결 상태 미리 확인
        try:
            from email_module.renewal_email import check_outlook_connection
            if not check_outlook_connection():
                messagebox.showerror("아웃룩 연결 오류", 
                    "아웃룩에 연결할 수 없습니다.\n\n아웃룩을 실행한 후 다시 시도해주세요.")
                return
        except Exception as e:
            # 연결 확인에 실패해도 계속 진행
            pass
        
        # 유통유무 체크
        if row.get('유통유무','') == '유통':
            messagebox.showwarning("경고", "유통 건은 직접 견적서를 작성해 주세요")
            return
        # Foundry 체크
        if row.get('대분류','') == 'Foundry':
            messagebox.showwarning("경고", "Foundry 제품은 영업사원이 직접 견적서를 작성해야 합니다")
            return
        # 만료일 안전 처리
        expiry_date = safe_parse_date(row['만료일'])
        if expiry_date is None:
            messagebox.showerror("오류", "만료일이 유효하지 않습니다. 데이터를 확인해주세요.")
            return
        
        sel_month = expiry_date.month
        client = row.get('고객사명','')
        same = [j for j,r in enumerate(self.rows)
                if safe_parse_date(r['만료일']) is not None and 
                safe_parse_date(r['만료일']).month == sel_month and 
                r.get('고객사명','')==client]
        if len(same) >= 2:
            dlg = tk.Toplevel(self.root); dlg.title("견적서 작성 방식 선택")
            tk.Label(dlg, text="2건 이상이 같은달에 있는 고객사 입니다. 어떻게 하시겠어요?", padx=20, pady=10).pack()
            def single_q():
                dlg.destroy()
                res_quote.create_and_send_quote(row, self.username)
                self.tree.set(iid, '액션', '메일+견적서')
                messagebox.showinfo("완료","견적서 작성 로그가 기록되었습니다.")
            def multi_q():
                dlg.destroy()
                lst = [self.rows[j] for j in same]
                res_quote_multi.create_and_send_quotes(lst, self.username)
                for j in same:
                    child_iid = self.tree.get_children()[j]
                    self.tree.set(child_iid, '액션', '메일+견적서')
                messagebox.showinfo("완료", f"{len(same)}건 견적서 작성 로그가 기록되었습니다.")
            frm = tk.Frame(dlg); frm.pack(pady=10)
            tk.Button(frm, text="1건만 작성", width=12, command=single_q).pack(side='left', padx=5)
            tk.Button(frm, text="2건 이상 작성", width=12, command=multi_q).pack(side='right', padx=5)
            dlg.transient(self.root); dlg.grab_set(); self.root.wait_window(dlg)
            return
        res_quote.create_and_send_quote(row, self.username)
        self.tree.set(iid, '액션', '메일+견적서')
        messagebox.showinfo("완료","견적서 작성 로그가 기록되었습니다.")

    def on_customer_search(self):
        """고객사명 콤보박스에서 엔터키로 검색"""
        entered_customer = self.customer_cb.get().strip()
        if entered_customer:
            # 디버깅용 로그 추가
            print(f"고객사 검색: '{entered_customer}', 전체 고객사 수: {len(self.all_customers)}")
            
            if not self.all_customers:
                print("경고: 고객사 목록이 비어있습니다!")
                return
                
            # 고객사명으로 검색 (대소문자 구분 없이)
            matching_customers = [c for c in self.all_customers if entered_customer.lower() in c.lower()]
            print(f"매칭되는 고객사: {matching_customers}")
            
            if matching_customers:
                # 검색 결과를 콤보박스에 설정
                self.customer_cb['values'] = ['전체'] + matching_customers
                # 첫 번째 매칭되는 고객사명을 설정
                self.customer_cb.set(matching_customers[0])
                # 드롭다운 리스트 표시
                self.customer_cb.event_generate('<Down>')
                self.customer_cb.event_generate('<Up>')
            else:
                # 매칭되는 고객사가 없으면 입력값 그대로 사용
                self.customer_cb.set(entered_customer)

    def update_customer_list(self):
        """고객사 필터 자동 완성"""
        val = self.customer_cb.get().strip()
        filtered = (
            self.all_customers
            if not val or val == '전체'
            else [c for c in self.all_customers if val.lower() in c.lower()]
        )
        self.customer_cb['values'] = ['전체'] + filtered

    def update_client_list(self):
        """거래처 필터 자동 완성"""
        val = self.client_cb.get().strip()
        filtered = (
            self.all_clients
            if not val or val == '전체'
            else [c for c in self.all_clients if val.lower() in c.lower()]
        )
        self.client_cb['values'] = ['전체'] + filtered

    def on_client_search(self):
        """거래처명 콤보박스에서 엔터키로 검색"""
        entered_client = self.client_cb.get().strip()
        if entered_client:
            # 디버깅용 로그 추가
            print(f"거래처 검색: '{entered_client}', 전체 거래처 수: {len(self.all_clients)}")
            
            if not self.all_clients:
                print("경고: 거래처 목록이 비어있습니다!")
                return
                
            # 거래처명으로 검색 (대소문자 구분 없이)
            matching_clients = [c for c in self.all_clients if entered_client.lower() in c.lower()]
            print(f"매칭되는 거래처: {matching_clients}")
            
            if matching_clients:
                # 검색 결과를 콤보박스에 설정 (모든 매칭 항목 표시)
                self.client_cb['values'] = ['전체'] + matching_clients
                # 입력값 그대로 유지 (첫 번째 항목으로 자동 설정하지 않음)
                self.client_cb.set(entered_client)
                # 드롭다운 리스트 표시
                self.client_cb.event_generate('<Down>')
                self.client_cb.event_generate('<Up>')
            else:
                # 매칭되는 거래처가 없으면 입력값 그대로 사용
                self.client_cb.set(entered_client)

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
        elif hasattr(date_value, 'month'):
            return date_value
        else:
            return None
    except (ValueError, TypeError):
        return None

def try_parse_date(series):
    # 1차: yyyy-mm-dd
    dt = pd.to_datetime(series, errors='coerce', format='%Y-%m-%d')
    if dt.notna().any():
        return dt
    # 2차: yyyy.mm.dd
    dt = pd.to_datetime(series, errors='coerce', format='%Y.%m.%d')
    if dt.notna().any():
        return dt
    # 3차: yyyy/mm/dd
    dt = pd.to_datetime(series, errors='coerce', format='%Y/%m/%d')
    if dt.notna().any():
        return dt
    # 4차: 포맷 없이(마지막 시도)
    return pd.to_datetime(series, errors='coerce')

if __name__ == '__main__':
    root = tk.Tk()
    RenewalMailerGUI(root, username="테스트유저")
    root.mainloop()
