import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from modules.renewal import renewal_search as rs
from modules.customer.customer_add import CustomerAddWindow
from core.data_loader import get_global_auth, safe_api_call
from integrations.lisa_logging import setup_logger
from utils.format_utils import safe_int_convert

# 모듈 로거 설정
logger = setup_logger('customer_info')

class CustomerInfoGUI:
    """고객사 관리 탭 내 임베디드 UI"""
    def __init__(self, parent_frame, username=None, data_loader=None, quote_gui=None):
        self.parent_frame = parent_frame
        self.username = username or '테스트유저'
        self.data_loader = data_loader
        self.quote_gui = quote_gui  # 견적서 탭 참조 저장
        
        # 기준 선택 변수 (기본값: 만료일)
        self.criteria_var = tk.StringVar(value="계산서 발행일")
        
        # 현재 선택된 기간 추적
        self.current_period = 'this_month'
        
        # 데이터 로더가 있으면 사용, 없으면 기존 방식 사용
        if self.data_loader:
            try:
                self.renewals_df = self.data_loader.get_renewal_data()
                
                # 데이터가 비어있으면 기존 방식으로 로드
                if self.renewals_df.empty:
                    print("고객사 관리: 데이터 로더에서 데이터를 가져올 수 없어 기존 방식으로 로드합니다.")
                    self._load_data_directly()
                else:
                    print(f"고객사 관리: 데이터 로더에서 데이터를 성공적으로 가져왔습니다: {len(self.renewals_df)}행")
                
                # 고객사 데이터는 별도로 로드
                try:
                    self.contacts_df = self.data_loader.get_customer_data()
                    if self.contacts_df.empty:
                        print("고객사 관리: 고객사 데이터가 비어있어 기본 데이터프레임을 생성합니다.")
                        self.contacts_df = pd.DataFrame()
                except Exception as e:
                    print(f"고객사 관리: 고객사 데이터 로드 오류 - {e}, 기본 데이터프레임을 생성합니다.")
                    self.contacts_df = pd.DataFrame()
            except Exception as e:
                print(f"고객사 관리: 데이터 로더 오류 - {e}, 기존 방식으로 로드합니다.")
                self._load_data_directly()
        else:
            # 기존 방식 (하위 호환성)
            print("고객사 관리: 기존 방식으로 데이터를 로드합니다.")
            self._load_data_directly()
        
        # 날짜 컬럼을 datetime으로 변환
        if '만료일' in self.renewals_df.columns:
            self.renewals_df['만료일'] = pd.to_datetime(self.renewals_df['만료일'], errors='coerce')
        if '계산서 발행일' in self.renewals_df.columns:
            self.renewals_df['계산서 발행일'] = pd.to_datetime(self.renewals_df['계산서 발행일'], errors='coerce')

        # 필터 옵션 리스트 (안전한 데이터 접근)
        self.all_customers = []
        if not self.renewals_df.empty and '고객사명' in self.renewals_df.columns:
            self.all_customers = sorted(self.renewals_df['고객사명'].dropna().unique().tolist())
            print(f"고객사 목록 로드 완료: {len(self.all_customers)}개 고객사")
            print(f"고객사 목록 샘플: {self.all_customers[:10]}")  # 처음 10개만 출력
        
        self.all_clients = []
        if not self.renewals_df.empty and '거래처명' in self.renewals_df.columns:
            self.all_clients = sorted(self.renewals_df['거래처명'].dropna().unique().tolist())
        
        self.all_categories = []
        if not self.renewals_df.empty and '대분류' in self.renewals_df.columns:
            self.all_categories = sorted(self.renewals_df['대분류'].dropna().unique().tolist())
        
        self.all_sales = []
        if not self.renewals_df.empty and '영업사원' in self.renewals_df.columns:
            self.all_sales = sorted(self.renewals_df['영업사원'].dropna().unique().tolist())

        # 천 단위 콤마 적용할 금액 컬럼 (입금액, 잔액 추가)
        self.cost_cols    = [
            '판매 단가','판매 합계','원가','원가계',
            '실 매입 단가','실 매입 합계','이익액','실 이익액',
            '입금액','잔액'
        ]
        # 숫자 우측 정렬할 컬럼 (금액 및 이익율)
        self.numeric_cols = self.cost_cols + ['이익율','실 이익율']

        self.sort_column  = None
        self.sort_reverse = False

        # 프레임 생성
        self.frame = tk.Frame(self.parent_frame)
        self.frame.pack(fill='both', expand=True)

        self.setup_widgets()
        self.load_filters()
        # 앱 실행 시 기본 "이번달" 표시
        self.load_period('this_month')
    
    def on_full_data_loaded(self):
        """전체 데이터 로딩 완료 시 호출되는 메서드"""
        if self.data_loader:
            try:
                # 데이터 로더에서 최신 데이터 가져오기
                self.renewals_df = self.data_loader.get_renewal_data()
                self.contacts_df = self.data_loader.get_customer_data()
                
                # 데이터가 비어있으면 기존 방식으로 로드
                if self.renewals_df.empty:
                    print("고객사 관리: 전체 데이터 로딩에서 데이터를 가져올 수 없어 기존 방식으로 로드합니다.")
                    self._load_data_directly()
                
                # 날짜 컬럼을 datetime으로 변환
                if '만료일' in self.renewals_df.columns:
                    self.renewals_df['만료일'] = pd.to_datetime(self.renewals_df['만료일'], errors='coerce')
                if '계산서 발행일' in self.renewals_df.columns:
                    self.renewals_df['계산서 발행일'] = pd.to_datetime(self.renewals_df['계산서 발행일'], errors='coerce')
                
                # 필터 옵션 업데이트
                self.all_customers = sorted(
                    self.renewals_df['고객사명'].dropna().unique().tolist()
                ) if '고객사명' in self.renewals_df.columns else []
                self.all_clients = sorted(
                    self.renewals_df['거래처명'].dropna().unique().tolist()
                ) if '거래처명' in self.renewals_df.columns else []
                self.all_categories = sorted(
                    self.renewals_df['대분류'].dropna().unique().tolist()
                ) if '대분류' in self.renewals_df.columns else []
                self.all_sales = sorted(
                    self.renewals_df['영업사원'].dropna().unique().tolist()
                ) if '영업사원' in self.renewals_df.columns else []
                
                # UI 업데이트
                self.load_filters()
                self.load_period(self.current_period)
                
                print(f"고객사 관리: 전체 데이터 로딩 완료")
                print(f"renewals_df 크기: {len(self.renewals_df)}행")
            except Exception as e:
                print(f"고객사 관리: 전체 데이터 로딩 오류 - {e}")
                # 에러 발생 시 기존 데이터 유지

    def on_criteria_changed(self):
        """조회 기준 변경 시 현재 기간으로 다시 필터링"""
        # 현재 선택된 기간을 다시 적용
        if hasattr(self, 'date_entry'):
            self.load_period(self.current_period)

    def _load_data_directly(self):
        """기존 방식으로 데이터 로드 (최적화된 방식)"""
        try:
            # 전역 인증 사용
            gc = get_global_auth()
            
            # ── 고객사 목록 로드 (renewal_list에서 고객사명별 중복 제거) ──
            def load_customer_data():
                sh = gc.open_by_key(rs.RENEWAL_LIST_SHEET_ID)
                ws = sh.worksheet(rs.RENEWAL_SHEET_NAME)
                return ws.get_all_values()
            
            all_values = safe_api_call(load_customer_data)
            if all_values:
                raw_header = [h.strip() for h in all_values[0]]
                valid_idx  = [i for i, h in enumerate(raw_header) if h]
                headers    = [raw_header[i] for i in valid_idx]
                rows       = [[row[i] for i in valid_idx] for row in all_values[1:]]
                df_all = pd.DataFrame(rows, columns=headers)
                df_all.columns = df_all.columns.str.strip()
                # 고객사명별로 한 행만 남기고(중복 제거), 고객사 정보 컬럼만 남김
                info_cols = [c for c in [
                    '고객사명','영업 담당자','담당자명','담당자','직함','이메일','연락처','핸드폰','주소','사업자번호','영문 회사명','영문 담당자명','영문 주소','비고'
                ] if c in df_all.columns]
                if '고객사명' in df_all.columns and info_cols:
                    self.contacts_df = df_all.drop_duplicates(subset='고객사명')[info_cols]
                elif info_cols:
                    self.contacts_df = pd.DataFrame(columns=info_cols)
                else:
                    self.contacts_df = pd.DataFrame()
            else:
                self.contacts_df = pd.DataFrame()
            self.contacts_df.columns = self.contacts_df.columns.str.strip()

            # ── 리뉴얼 목록 로드 ──
            def load_renewal_data():
                sh = gc.open_by_key(rs.RENEWAL_LIST_SHEET_ID)
                ws = sh.worksheet(rs.RENEWAL_SHEET_NAME)
                return ws.get_all_records()
            
            renewal_data = safe_api_call(load_renewal_data)
            if renewal_data:
                self.renewals_df = pd.DataFrame(renewal_data)
                self.renewals_df.columns = self.renewals_df.columns.str.strip()
                
                # 날짜 컬럼을 datetime으로 변환
                if '만료일' in self.renewals_df.columns:
                    self.renewals_df['만료일'] = pd.to_datetime(self.renewals_df['만료일'], errors='coerce')
                if '계산서 발행일' in self.renewals_df.columns:
                    self.renewals_df['계산서 발행일'] = pd.to_datetime(self.renewals_df['계산서 발행일'], errors='coerce')
            else:
                self.renewals_df = pd.DataFrame()
            
            # 필터 옵션 리스트 업데이트 (renewals_df에서 가져오기)
            self.all_customers = []
            if not self.renewals_df.empty and '고객사명' in self.renewals_df.columns:
                self.all_customers = sorted(self.renewals_df['고객사명'].dropna().unique().tolist())
                print(f"고객사 목록 로드 완료: {len(self.all_customers)}개 고객사")
                
            self.all_clients = []
            if not self.renewals_df.empty and '거래처명' in self.renewals_df.columns:
                self.all_clients = sorted(self.renewals_df['거래처명'].dropna().unique().tolist())
                
            self.all_categories = []
            if not self.renewals_df.empty and '대분류' in self.renewals_df.columns:
                self.all_categories = sorted(self.renewals_df['대분류'].dropna().unique().tolist())
                
        except Exception as e:
            print(f"데이터 로드 오류: {e}")
            self.contacts_df = pd.DataFrame()
            self.renewals_df = pd.DataFrame()

    def setup_widgets(self):
        # 상단 헤더
        header_frame = tk.Frame(self.frame, bg="#F7F9FB", height=60)
        header_frame.pack(fill='x', pady=(0, 20))
        header_frame.pack_propagate(False)
        
        # 헤더 제목
        title_label = tk.Label(header_frame, text="고객사 관리", 
                              fg="#1F2937", bg="#F7F9FB")
        title_label.pack(side='left', padx=20, pady=15)
        
        # 사용자 정보 (우측)
        user_label = tk.Label(header_frame, text=f"👤 {self.username}", 
                             fg="#6B7280", bg="#F7F9FB")
        user_label.pack(side='right', padx=20, pady=15)

        # ── 필터 카드 영역 ──
        filter_section = tk.Frame(self.frame, bg="#F7F9FB")
        filter_section.pack(fill='x', padx=20, pady=(0, 20))

        # 첫 번째 카드: 기준 선택 및 날짜
        card1 = tk.Frame(filter_section, bg="white", relief="solid", bd=1)
        card1.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        # 카드1 헤더
        card1_header = tk.Frame(card1, bg="#F8FAFC", height=40)
        card1_header.pack(fill='x')
        card1_header.pack_propagate(False)
        tk.Label(card1_header, text="📅 조회 기준", 
                fg="#374151", bg="#F8FAFC").pack(side='left', padx=15, pady=10)
        
        # 카드1 내용
        card1_content = tk.Frame(card1, bg="white")
        card1_content.pack(fill='both', expand=True, padx=15, pady=15)
        
        # 기준 선택
        criteria_frame = tk.Frame(card1_content, bg="white")
        criteria_frame.pack(fill='x', pady=(0, 10))
        tk.Label(criteria_frame, text="조회 기준:", 
                fg="#374151", bg="white").pack(side='left')
        tk.Radiobutton(criteria_frame, text="계산서 발행일", variable=self.criteria_var, 
                      value="계산서 발행일", command=self.on_criteria_changed,
                      fg="#374151", bg="white").pack(side='left', padx=5)
        tk.Radiobutton(criteria_frame, text="만료일", variable=self.criteria_var, 
                      value="만료일", command=self.on_criteria_changed,
                      fg="#374151", bg="white").pack(side='left', padx=5)
        
        # 날짜 선택
        date_frame = tk.Frame(card1_content, bg="white")
        date_frame.pack(fill='x', pady=(0, 10))
        tk.Label(date_frame, text="조회 기준일:", 
                fg="#374151", bg="white").pack(side='left')
        from tkcalendar import DateEntry
        from datetime import datetime
        self.date_entry = DateEntry(
            date_frame,
            width=12,
            date_pattern='yyyy-mm-dd'
        )
        self.date_entry.set_date(datetime.today())
        self.date_entry.pack(side='left', padx=5)
        # 날짜 선택 시 현재 기간 유지
        self.date_entry.bind(
            "<<DateEntrySelected>>",
            lambda e: self.load_period(self.current_period)
        )
        
        # 기간 버튼들 - 2줄로 분리
        # 첫 번째 줄: 이번달, 다음달, 미래 1년
        first_row_frame = tk.Frame(card1_content, bg="white")
        first_row_frame.pack(fill='x', pady=(0, 5))
        
        first_row_periods = [
            ("이번달",   'this_month'),
            ("다음달",   'next_month'),
            ("미래 1년", 'one_year')
        ]
        for text, mode in first_row_periods:
            btn = tk.Button(
                first_row_frame, text=text,
                command=lambda m=mode: self.load_period(m),
                bg="#3B82F6", fg="white",
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
                                 command=lambda: self.load_period('past_year'),
                                 bg="#3B82F6", fg="white",
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
        self.customer_cb = ttk.Combobox(
            row1,
            state='normal',
            width=20,
            postcommand=self.update_customer_list
        )
        self.customer_cb.pack(side='left', padx=5)
        self.customer_cb.bind(
            '<<ComboboxSelected>>',
            lambda e: self.on_search()
        )
        self.customer_cb.bind('<Return>', lambda e: self.on_customer_search())
        
        # 고객사명 엔터키 힌트 제거

        tk.Label(row1, text="거래처명:", font=("맑은 고딕", 10), 
                fg="#374151", bg="white").pack(side='left', padx=(20,0))
        self.client_cb = ttk.Combobox(
            row1,
            state='normal',
            width=20,
            postcommand=self.update_client_list
        )
        self.client_cb.pack(side='left', padx=5)
        self.client_cb.bind(
            '<<ComboboxSelected>>',
            lambda e: self.on_search()
        )
        self.client_cb.bind('<Return>', lambda e: self.on_client_search())
        
        # 거래처명 엔터키 힌트 제거

        # 두 번째 행: 대분류, 영업사원
        row2 = tk.Frame(card2_content, bg="white")
        row2.pack(fill='x', pady=(0, 10))
        
        tk.Label(row2, text="대분류:", font=("맑은 고딕", 10), 
                fg="#374151", bg="white").pack(side='left')
        self.category_cb = ttk.Combobox(
            row2,
            state='normal',
            width=20,
            postcommand=self.update_category_list
        )
        self.category_cb.pack(side='left', padx=5)
        self.category_cb.bind(
            '<<ComboboxSelected>>',
            lambda e: self.on_search()
        )
        self.category_cb.bind('<Return>', lambda e: self.on_category_search())
        
        # 대분류 엔터키 힌트 제거

        tk.Label(row2, text="영업사원:", font=("맑은 고딕", 10), 
                fg="#374151", bg="white").pack(side='left', padx=(20,0))
        self.sales_cb = ttk.Combobox(
            row2,
            state='normal',
            width=20,
            postcommand=self.update_sales_list
        )
        self.sales_cb.pack(side='left', padx=5)
        self.sales_cb.bind(
            '<<ComboboxSelected>>',
            lambda e: self.on_search()
        )
        self.sales_cb.bind('<Return>', lambda e: self.on_sales_search())
        
        # 영업사원 엔터키 힌트 제거

        # 세 번째 행: 검색 버튼만
        row3 = tk.Frame(card2_content, bg="white")
        row3.pack(fill='x')
        
        search_btn = tk.Button(row3, text="🔍 검색", command=self.on_search,
                              font=("맑은 고딕", 10, "bold"), bg="#10B981", fg="white",
                              relief="flat", padx=15, pady=4)
        search_btn.pack(side='left')
        search_btn.bind('<Enter>', lambda e: search_btn.configure(bg="#059669"))
        search_btn.bind('<Leave>', lambda e: search_btn.configure(bg="#10B981"))

        # 세 번째 카드: 액션 프레임 (필터 오른쪽에 배치)
        card3 = tk.Frame(filter_section, bg="white", relief="solid", bd=1)
        card3.pack(side='left', fill='both', expand=True)
        
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
        
        # 신규 고객사 등록 버튼
        add_btn = tk.Button(action_buttons, text="➕ 신규 고객사 등록",
                           command=lambda: CustomerAddWindow(self.frame, self.reload_customers),
                           font=("맑은 고딕", 10, "bold"), bg="#3B82F6", fg="white",
                           relief="flat", padx=15, pady=4)
        add_btn.pack(fill='x', pady=(0, 5))
        add_btn.bind('<Enter>', lambda e: add_btn.configure(bg="#2563EB"))
        add_btn.bind('<Leave>', lambda e: add_btn.configure(bg="#3B82F6"))
        
        # 견적 작성(기존) 버튼
        quote_btn = tk.Button(action_buttons, text="📋 견적 작성(기존)", command=self.open_quote_simple,
                             font=("맑은 고딕", 10, "bold"), bg="#8B5CF6", fg="white",
                             relief="flat", padx=15, pady=4)
        quote_btn.pack(fill='x')
        quote_btn.bind('<Enter>', lambda e: quote_btn.configure(bg="#7C3AED"))
        quote_btn.bind('<Leave>', lambda e: quote_btn.configure(bg="#8B5CF6"))



        # ── 구매 정보 테이블 섹션 ──
        purchase_section = tk.Frame(self.frame, bg="#F7F9FB")
        purchase_section.pack(fill='both', expand=True, padx=20, pady=(0, 20))
        
        # 테이블 헤더
        purchase_header = tk.Frame(purchase_section, bg="white", borderwidth=1, relief="solid", border=1)
        purchase_header.pack(fill='x', pady=(0, 1))
        tk.Label(purchase_header, text="💰 구매 정보", font=("맑은 고딕", 14, "bold"), 
                fg="#1F2937", bg="white").pack(side='left', padx=15, pady=10)
        
        # 테이블 컨테이너
        purchase_container = tk.Frame(purchase_section, bg="white", borderwidth=1, relief="solid", border=1)
        purchase_container.pack(fill='both', expand=True)
        
        self.purchase_cols = list(self.renewals_df.columns[:19])  # A~S열
        self.purchase_tree = ttk.Treeview(purchase_container,
                                          columns=self.purchase_cols,
                                          show='headings',
                                          selectmode='browse')
        vsb = ttk.Scrollbar(purchase_container, orient='vertical',
                             command=self.purchase_tree.yview)
        hsb = ttk.Scrollbar(purchase_container, orient='horizontal',
                             command=self.purchase_tree.xview)
        self.purchase_tree.configure(yscrollcommand=vsb.set,
                                     xscrollcommand=hsb.set)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        self.purchase_tree.pack(fill='both', expand=True, padx=1, pady=1)

        # 컬럼 크기 조절 기능 추가
        from ui.resizable_treeview import ResizableTreeview
        self.resizable_tree = ResizableTreeview(self.purchase_tree, "customer_purchase_table")

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
        self.purchase_tree.bind("<Shift-MouseWheel>", fast_horizontal_scroll)

        for col in self.purchase_cols:
            anchor = 'e' if col in self.numeric_cols else 'center'
            self.purchase_tree.heading(col, text=col,
                command=lambda c=col: self.sort_purchase(c))
            self.purchase_tree.column(col, width=100,
                                     anchor=anchor, stretch=False)
        self.purchase_tree.bind('<Double-1>', self.on_tree_double_click)

    def load_filters(self):
        self.customer_cb['values'] = ['전체'] + self.all_customers
        self.customer_cb.set('전체')
        self.client_cb['values']   = ['전체'] + self.all_clients
        self.client_cb.set('전체')
        self.category_cb['values'] = ['전체'] + self.all_categories
        self.category_cb.set('전체')
        self.sales_cb['values'] = ['전체'] + self.all_sales
        self.sales_cb.set('전체')
        
        # 초기 로딩 시 전체 데이터 표시
        if hasattr(self, 'renewals_df') and len(self.renewals_df) > 0:
            self.filtered_df = self.renewals_df.copy()
            self.load_data()
            print(f"초기 로딩: 전체 데이터 {len(self.renewals_df)}행 표시")
        
        # 앱 실행 시 기본 "이번달" 표시
        self.load_period('this_month')

    def load_period(self, mode):
        """기간별 데이터 필터링"""
        from datetime import datetime, timedelta
        import calendar
        
        # 현재 선택된 기간 업데이트
        self.current_period = mode
        
        selected_date = self.date_entry.get_date()
        
        if mode == 'this_month':
            # 이번달 (1일 ~ 말일)
            start_date = selected_date.replace(day=1)
            _, last_day = calendar.monthrange(selected_date.year, selected_date.month)
            end_date = selected_date.replace(day=last_day)
        elif mode == 'next_month':
            # 다음달
            if selected_date.month == 12:
                start_date = selected_date.replace(year=selected_date.year + 1, month=1, day=1)
            else:
                start_date = selected_date.replace(month=selected_date.month + 1, day=1)
            _, last_day = calendar.monthrange(start_date.year, start_date.month)
            end_date = start_date.replace(day=last_day)
        elif mode == 'one_year':
            # 미래 1년
            start_date = selected_date
            end_date = selected_date + timedelta(days=365)
        elif mode == 'past_year':
            # 과거 1년
            start_date = selected_date - timedelta(days=365)
            end_date = selected_date
        else:
            return
        
        # 선택된 기준에 따라 필터링
        start_dt = pd.to_datetime(start_date)
        end_dt = pd.to_datetime(end_date)
        
        print(f"기간 필터링: {mode}")
        print(f"기준일: {selected_date}")
        print(f"시작일: {start_date}, 종료일: {end_date}")
        print(f"조회 기준: {self.criteria_var.get()}")
        print(f"전체 데이터 수: {len(self.renewals_df)}")
        
        # 원본 데이터 샘플 확인
        if len(self.renewals_df) > 0:
            print(f"원본 데이터 컬럼: {list(self.renewals_df.columns)}")
            if '계산서 발행일' in self.renewals_df.columns:
                sample_dates = self.renewals_df['계산서 발행일'].dropna().head(5).tolist()
                print(f"원본 계산서 발행일 샘플: {sample_dates}")
            if '만료일' in self.renewals_df.columns:
                sample_dates = self.renewals_df['만료일'].dropna().head(5).tolist()
                print(f"원본 만료일 샘플: {sample_dates}")
        
        # 원본 데이터에서 필터링 (원본 데이터 보존)
        # 원본 데이터를 별도로 저장
        original_df = self.renewals_df.copy()
        
        if self.criteria_var.get() == "계산서 발행일" and '계산서 발행일' in original_df.columns:
            # NaT 값 제외하고 필터링
            mask = (
                (original_df['계산서 발행일'].notna()) &
                (original_df['계산서 발행일'] >= start_dt) & 
                (original_df['계산서 발행일'] <= end_dt)
            )
            filtered_df = original_df[mask].copy()
            print(f"계산서 발행일 기준 필터링 결과: {len(filtered_df)}행")
            
            # 필터링 후 데이터 샘플 확인
            if len(filtered_df) > 0:
                sample_dates = filtered_df['계산서 발행일'].head(5).tolist()
                print(f"필터링 후 계산서 발행일 샘플: {sample_dates}")
        elif self.criteria_var.get() == "만료일" and '만료일' in original_df.columns:
            # NaT 값 제외하고 필터링
            mask = (
                (original_df['만료일'].notna()) &
                (original_df['만료일'] >= start_dt) & 
                (original_df['만료일'] <= end_dt)
            )
            filtered_df = original_df[mask].copy()
            print(f"만료일 기준 필터링 결과: {len(filtered_df)}행")
        else:
            filtered_df = original_df.copy()
            print(f"필터링 없음: {len(filtered_df)}행")
        
        # 필터링된 데이터를 표시용으로 저장
        self.filtered_df = filtered_df
        
        # 데이터가 없어도 UI는 정상적으로 표시되도록 load_data 호출
        self.load_data()
        
        # 데이터가 없을 때 사용자에게 알림 (선택사항)
        if len(filtered_df) == 0:
            print(f"해당 기간({start_date.strftime('%Y-%m-%d')} ~ {end_date.strftime('%Y-%m-%d')})에 데이터가 없습니다.")

    def refresh_data(self):
        """데이터 새로고침"""
        try:
            print("고객사 관리: 데이터 새로고침 시작")
            
            # 데이터 로더가 있으면 전체 데이터 다시 로드
            if self.data_loader:
                # 현재 데이터 새로고침
                self.contacts_df = self.data_loader.get_customer_data()
                self.renewals_df = self.data_loader.get_renewal_data()
                
                # 날짜 컬럼을 datetime으로 변환
                if '만료일' in self.renewals_df.columns:
                    self.renewals_df['만료일'] = pd.to_datetime(self.renewals_df['만료일'], errors='coerce')
                if '계산서 발행일' in self.renewals_df.columns:
                    self.renewals_df['계산서 발행일'] = pd.to_datetime(self.renewals_df['계산서 발행일'], errors='coerce')
                
                # 필터 옵션 리스트 업데이트
                self.all_customers = []
                if not self.renewals_df.empty and '고객사명' in self.renewals_df.columns:
                    self.all_customers = sorted(self.renewals_df['고객사명'].dropna().unique().tolist())
                
                self.all_clients = []
                if not self.renewals_df.empty and '거래처명' in self.renewals_df.columns:
                    self.all_clients = sorted(self.renewals_df['거래처명'].dropna().unique().tolist())
                
                self.all_categories = []
                if not self.renewals_df.empty and '대분류' in self.renewals_df.columns:
                    self.all_categories = sorted(self.renewals_df['대분류'].dropna().unique().tolist())
                
                # 필터 옵션 다시 로드
                self.load_filters()
                # 현재 표시 중인 데이터 다시 로드 (데이터가 없어도 UI는 정상적으로 표시)
                self.load_period(self.current_period)
                
                print(f"고객사 관리: 데이터 새로고침 완료 - 고객사: {len(self.contacts_df)}행, 리뉴얼: {len(self.renewals_df)}행")
                
                # 데이터가 없을 때 사용자에게 알림 (선택사항)
                if len(self.renewals_df) == 0:
                    print("고객사 관리: 리뉴얼 데이터가 없습니다.")
                if len(self.contacts_df) == 0:
                    print("고객사 관리: 고객사 데이터가 없습니다.")
            else:
                # 데이터 로더가 없으면 기존 방식으로 로드
                self._load_data_directly()
                self.load_filters()
                self.load_period(self.current_period)
                print("고객사 관리: 기존 방식으로 데이터 새로고침 완료")
                
        except Exception as e:
            print(f"고객사 관리: 데이터 새로고침 오류 - {e}")

    def load_all_data(self):
        """전체 데이터 로드 (기간 제한 없음)"""
        print("고객사 관리: 전체 데이터 로드 버튼 클릭")
        try:
            # 필터 필드 초기화
            self.customer_cb.set('전체')
            self.client_cb.set('전체')
            self.category_cb.set('전체')
            
            # 데이터가 비어있으면 먼저 직접 로드 시도
            if self.renewals_df is None or self.renewals_df.empty:
                print("고객사 관리: 데이터가 없어 직접 로드 시도")
                self._load_data_directly()
            
            # 전체 데이터를 트리뷰에 표시
            if self.renewals_df is not None and not self.renewals_df.empty:
                print(f"고객사 관리: 전체 데이터 로드: {len(self.renewals_df)}행")
                self.filtered_df = self.renewals_df.copy()
                self.load_data()
                messagebox.showinfo("완료", f"전체 데이터 {len(self.renewals_df)}행을 로드했습니다.")
            else:
                print("고객사 관리: 로드할 데이터가 없습니다.")
                messagebox.showwarning("경고", "로드할 데이터가 없습니다. 네트워크 연결을 확인해주세요.")
        except Exception as e:
            print(f"고객사 관리: 전체 데이터 로드 오류: {e}")
            messagebox.showerror("오류", f"전체 데이터 로드 중 오류가 발생했습니다:\n{e}")

    def load_data(self):
        """필터링된 데이터를 테이블에 표시"""
        # 기존 데이터 삭제
        for item in self.purchase_tree.get_children():
            self.purchase_tree.delete(item)
        
        # 필터링된 데이터 표시 (filtered_df 사용)
        if hasattr(self, 'filtered_df'):
            df = self.filtered_df
            print(f"load_data: filtered_df 사용, {len(df)}행")
        else:
            df = self.renewals_df
            print(f"load_data: renewals_df 사용, {len(df)}행")
        
        if len(df) == 0:
            print("경고: 표시할 데이터가 없습니다!")
            # 데이터가 없어도 UI는 정상적으로 표시되도록 함
            # 빈 테이블 상태로 유지
            return
        
        # 데이터 샘플 확인 (디버깅용)
        if len(df) > 0:
            print(f"데이터 컬럼: {list(df.columns)}")
            print(f"첫 번째 행 데이터: {df.iloc[0].to_dict()}")
        
        inserted_count = 0
        for _, row in df.iterrows():
            values = []
            for col in self.purchase_cols:
                val = row.get(col, "")
                # 계산서 발행일과 만료일은 날짜만 표시
                if col in ['계산서 발행일', '만료일'] and val:
                    if isinstance(val, str):
                        # 시간 부분이 포함된 경우 날짜만 추출
                        if ' ' in val:
                            val = val.split(' ')[0]
                    elif hasattr(val, 'strftime'):
                        # NaT 값 체크
                        try:
                            if pd.notna(val):  # NaT가 아닌 경우에만 strftime 호출
                                val = val.strftime('%Y-%m-%d')
                            else:
                                val = ""  # NaT인 경우 빈 문자열로 설정
                        except:
                            val = ""  # 예외 발생 시 빈 문자열로 설정
                elif col in self.cost_cols and val:
                    try:
                        val = f"{safe_int_convert(val):,}"
                    except:
                        pass
                values.append(val)
            self.purchase_tree.insert("", "end", values=values)
            inserted_count += 1
        
        print(f"테이블에 {inserted_count}행 삽입 완료")
        
        # 첫 번째 행의 데이터 샘플 출력 (디버깅용)
        if inserted_count > 0:
            first_row = df.iloc[0]
            print(f"첫 번째 행 샘플: 고객사명='{first_row.get('고객사명', 'N/A')}', 거래처명='{first_row.get('거래처명', 'N/A')}', 계산서발행일='{first_row.get('계산서 발행일', 'N/A')}'")
        else:
            print("데이터가 없어서 테이블에 아무것도 삽입되지 않았습니다.")

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

    def update_category_list(self):
        """대분류 필터 자동 완성"""
        val = self.category_cb.get().strip()
        filtered = (
            self.all_categories
            if not val or val == '전체'
            else [c for c in self.all_categories if val.lower() in c.lower()]
        )
        self.category_cb['values'] = ['전체'] + filtered

    def update_sales_list(self):
        """영업사원 필터 자동 완성"""
        val = self.sales_cb.get().strip()
        filtered = (
            self.all_sales
            if not val or val == '전체'
            else [c for c in self.all_sales if val.lower() in c.lower()]
        )
        self.sales_cb['values'] = ['전체'] + filtered

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

    def on_sales_search(self):
        """영업사원 콤보박스에서 엔터키로 검색"""
        entered_sales = self.sales_cb.get().strip()
        if entered_sales:
            # 디버깅용 로그 추가
            print(f"영업사원 검색: '{entered_sales}', 전체 영업사원 수: {len(self.all_sales)}")
            
            if not self.all_sales:
                print("경고: 영업사원 목록이 비어있습니다!")
                return
                
            # 영업사원으로 검색 (대소문자 구분 없이)
            matching_sales = [c for c in self.all_sales if entered_sales.lower() in c.lower()]
            print(f"매칭되는 영업사원: {matching_sales}")
            
            if matching_sales:
                # 검색 결과를 콤보박스에 설정 (모든 매칭 항목 표시)
                self.sales_cb['values'] = ['전체'] + matching_sales
                # 입력값 그대로 유지 (첫 번째 항목으로 자동 설정하지 않음)
                self.sales_cb.set(entered_sales)
                # 드롭다운 리스트 표시
                self.sales_cb.event_generate('<Down>')
                self.sales_cb.event_generate('<Up>')
            else:
                # 매칭되는 영업사원이 없으면 입력값 그대로 사용
                self.sales_cb.set(entered_sales)

    def on_search(self):
        """고객사명, 거래처명 기준 데이터 필터링 및 UI 갱신"""
        # 위젯이 파괴되었는지 확인
        try:
            if not self.frame.winfo_exists():
                return
        except:
            return
            
        sel_cust   = self.customer_cb.get().strip()
        sel_client = self.client_cb.get().strip()
        sel_cat    = self.category_cb.get().strip()
        sel_sales  = self.sales_cb.get().strip()
        
        print(f"검색 조건: 고객사명='{sel_cust}', 거래처명='{sel_client}', 대분류='{sel_cat}', 영업사원='{sel_sales}'")
        
        # 구매정보 필터링 (조회 기간 필터링된 데이터 사용)
        if hasattr(self, 'filtered_df') and len(self.filtered_df) > 0:
            df_pur = self.filtered_df.copy()
            print(f"기간 필터링된 데이터 사용: {len(df_pur)}행")
        else:
            df_pur = self.renewals_df.copy()
            print(f"전체 데이터 사용: {len(df_pur)}행")
            
        if df_pur.empty:
            self.load_data()
            return
            
        print(f"필터 적용 시작: 원본 데이터 {len(df_pur)}행")
        
        try:
            # 대분류 필터링
            if sel_cat and sel_cat != '전체' and '대분류' in df_pur.columns:
                before_count = len(df_pur)
                df_pur = df_pur[df_pur['대분류'].str.contains(sel_cat, na=False, case=False, regex=False)]
                print(f"대분류 '{sel_cat}' 필터링: {before_count}행 → {len(df_pur)}행")
                
            # 고객사명 필터링
            if sel_cust and sel_cust != '전체' and '고객사명' in df_pur.columns:
                before_count = len(df_pur)
                # 대소문자 구분 없이 검색하고, NaN 값 처리
                df_pur = df_pur[df_pur['고객사명'].str.contains(sel_cust, na=False, case=False, regex=False)]
                print(f"고객사명 '{sel_cust}' 필터링: {before_count}행 → {len(df_pur)}행")
                if len(df_pur) == 0:
                    print(f"경고: 고객사명 '{sel_cust}'에 해당하는 데이터가 없습니다.")
                    print(f"사용 가능한 고객사명 샘플: {df_pur['고객사명'].dropna().unique()[:5].tolist()}")
                    
            # 거래처명 필터링
            if sel_client and sel_client != '전체' and '거래처명' in df_pur.columns:
                before_count = len(df_pur)
                # 대소문자 구분 없이 검색하고, NaN 값 처리
                df_pur = df_pur[df_pur['거래처명'].str.contains(sel_client, na=False, case=False, regex=False)]
                print(f"거래처명 '{sel_client}' 필터링: {before_count}행 → {len(df_pur)}행")
                if len(df_pur) == 0:
                    print(f"경고: 거래처명 '{sel_client}'에 해당하는 데이터가 없습니다.")
                    
            # 영업사원 필터링
            if sel_sales and sel_sales != '전체' and '영업사원' in df_pur.columns:
                before_count = len(df_pur)
                # 대소문자 구분 없이 검색하고, NaN 값 처리
                df_pur = df_pur[df_pur['영업사원'].str.contains(sel_sales, na=False, case=False, regex=False)]
                print(f"영업사원 '{sel_sales}' 필터링: {before_count}행 → {len(df_pur)}행")
                if len(df_pur) == 0:
                    print(f"경고: 영업사원 '{sel_sales}'에 해당하는 데이터가 없습니다.")
                    
        except Exception as e:
            print(f"필터 적용 오류: {e}")
            import traceback
            traceback.print_exc()

        print(f"필터 적용 완료: 최종 {len(df_pur)}행")
        
        # 실 매입 단가 우선 교체
        if '실 매입 단가' in df_pur.columns:
            mask = (df_pur['실 매입 단가'].notna()
                    & (df_pur['실 매입 단가'] != df_pur['원가']))
            
            # 값 추출 및 타입 변환
            cols_src = ['실 매입 단가','실 매입 합계','실 이익액','실 이익율']
            cols_dst = ['원가','원가계','이익액','이익율']
            
            values = df_pur.loc[mask, cols_src].copy()
            for col in values.columns:
                # int로 변환, 변환 불가시 NaN
                values[col] = pd.to_numeric(values[col], errors='coerce')
            
            # 대상 컬럼도 float으로 변환
            for col in cols_dst:
                df_pur[col] = pd.to_numeric(df_pur[col], errors='coerce')
            
            df_pur.loc[mask, cols_dst] = values.values

        # 정렬 유지
        if self.sort_column:
            df_pur = df_pur.sort_values(
                by=self.sort_column,
                ascending=not self.sort_reverse,
                key=lambda x: x.fillna('') if x.dtype==object else x
            )

        # 필터링된 데이터를 filtered_df에 저장하고 load_data 호출
        self.filtered_df = df_pur.copy()
        
        # load_data 호출하여 트리뷰 갱신 (데이터가 없어도 UI는 정상적으로 표시)
        self.load_data()

    def sort_purchase(self, col):
        # 위젯이 파괴되었는지 확인
        try:
            if not self.frame.winfo_exists():
                return
        except:
            return
        
        # 필터링된 데이터 사용 (전체 데이터가 아닌)
        if hasattr(self, 'filtered_df') and len(self.filtered_df) > 0:
            df = self.filtered_df.copy()
        else:
            # filtered_df가 없으면 전체 데이터 사용
            df = self.renewals_df.copy()

        rev = (self.sort_column==col) and not self.sort_reverse
        df_sorted = df.sort_values(
            by=col,
            ascending=not rev,
            key=lambda x: x.fillna('') if x.dtype==object else x
        )

        # 실 매입 단가 우선 교체
        if '실 매입 단가' in df_sorted.columns:
            mask = (df_sorted['실 매입 단가'].notna()
                    & (df_sorted['실 매입 단가'] != df_sorted['원가']))
            
            # 값 추출 및 타입 변환
            cols_src = ['실 매입 단가','실 매입 합계','실 이익액','실 이익율']
            cols_dst = ['원가','원가계','이익액','이익율']
            
            values = df_sorted.loc[mask, cols_src].copy()
            for col2 in values.columns:
                # int로 변환, 변환 불가시 NaN
                values[col2] = pd.to_numeric(values[col2], errors='coerce')
            
            # 대상 컬럼도 float으로 변환
            for col2 in cols_dst:
                df_sorted[col2] = pd.to_numeric(df_sorted[col2], errors='coerce')
            
            df_sorted.loc[mask, cols_dst] = values.values

        try:
            if self.purchase_tree.winfo_exists():
                self.purchase_tree.delete(*self.purchase_tree.get_children())
                for _, row in df_sorted.iterrows():
                    vals = []
                    for c in self.purchase_cols:
                        v = row.get(c, "")
                        if c in self.cost_cols:
                            try:    v = f"{safe_int_convert(v):,}"
                            except: pass
                        elif c in ['이익율','실 이익율']:
                            if v in [None, '', 'NaN', 'nan'] or (isinstance(v, float) and pd.isna(v)):
                                v = ''
                            else:
                                try:
                                    v = f"{float(v):.2f}"
                                except:
                                    pass
                        vals.append(v)
                    self.purchase_tree.insert('', 'end', values=vals)
        except:
            pass

        self.sort_column  = col
        self.sort_reverse = rev

    def on_tree_double_click(self, event):
        item = self.purchase_tree.identify_row(event.y)
        if not item:
            return
        vals = self.purchase_tree.item(item, 'values')
        if not vals:
            return
        # 고객사명 컬럼 인덱스 찾기
        cust_col = None
        for i, col in enumerate(self.purchase_cols):
            if col == '고객사명':
                cust_col = i
                break
        if cust_col is None:
            return
        customer_name = vals[cust_col]
        if not customer_name:
            return
        try:
            # 새로운 고객사 실적 다이얼로그 호출
            from customer_performance import CustomerPerformanceDialog
            CustomerPerformanceDialog(self.frame, customer_name, self.data_loader)
        except Exception as e:
            import traceback
            traceback.print_exc()
            tk.messagebox.showerror('오류', f'고객사 실적 창을 여는 데 실패했습니다:\n{e}')

    def reload_customers(self):
        """고객사 목록 재로딩"""
        try:
            if self.data_loader:
                # 데이터 로더에서 최신 데이터 가져오기
                self.contacts_df = self.data_loader.get_customer_data()
                
                if self.contacts_df.empty:
                    print("고객사 관리: 재로딩에서 데이터를 가져올 수 없어 기존 방식으로 로드합니다.")
                    self._load_data_directly()
            else:
                # 기존 방식으로 직접 로드
                self._load_data_directly()
                
            # 고객사 목록 업데이트
            self.all_customers = sorted(self.renewals_df['고객사명'].dropna().unique().tolist()) if '고객사명' in self.renewals_df.columns else []
            self.load_filters()
            
        except Exception as e:
            print(f"고객사 목록 재로딩 오류: {e}")
            self.contacts_df = pd.DataFrame()
            self.all_customers = []

    def open_quote_simple(self):
        print("견적 작성(기존) 버튼 클릭됨")
        # 구매정보 트리뷰에서 선택된 행의 정보를 가져와 quote_simple.py로 전달
        selected = self.purchase_tree.selection()
        if not selected:
            print("선택된 행이 없음")
            tk.messagebox.showinfo('알림', '구매정보에서 견적을 작성할 행을 선택하세요.')
            return
        # 1. 선택된 행의 계산서 발행일, 고객사명, 거래처명 추출
        vals = self.purchase_tree.item(selected[0], 'values')
        print(f"선택된 행 데이터: {vals}")
        print(f"구매정보 컬럼: {self.purchase_cols}")
        
        bill_date_idx = self.purchase_cols.index('계산서 발행일')
        customer_idx = self.purchase_cols.index('고객사명')
        client_idx = self.purchase_cols.index('거래처명')
        
        bill_date = vals[bill_date_idx]
        customer_name = vals[customer_idx]
        client_name = vals[client_idx]
        
        print(f"추출된 데이터: 계산서발행일='{bill_date}' (타입: {type(bill_date)}), 고객사명='{customer_name}' (타입: {type(customer_name)}), 거래처명='{client_name}' (타입: {type(client_name)})")

        # 2. 현재 구매정보 DataFrame에서 같은 계산서 발행일, 고객사명, 거래처명의 모든 행 필터링
        df = self.renewals_df.copy()
        print(f"전체 DataFrame 크기: {len(df)}")
        
        sel_cust = self.customer_cb.get().strip()
        sel_client = self.client_cb.get().strip()
        if sel_cust and sel_cust != '전체':
            df = df[df['고객사명'].str.contains(sel_cust, na=False, case=False, regex=False)]
            print(f"고객사 필터링 후 크기: {len(df)}")
        if sel_client and sel_client != '전체':
            df = df[df['거래처명'].str.contains(sel_client, na=False, case=False, regex=False)]
            print(f"거래처 필터링 후 크기: {len(df)}")
        
        # 디버깅: DataFrame의 처음 몇 행 출력
        print("DataFrame 처음 3행:")
        for i, (_, row) in enumerate(df.head(3).iterrows()):
            print(f"  행 {i}: 계산서발행일='{row['계산서 발행일']}', 고객사명='{row['고객사명']}', 거래처명='{row['거래처명']}'")
        
        # 선택된 날짜 형식 정규화 (시간 부분 제거)
        normalized_bill_date = str(bill_date)
        if ' ' in normalized_bill_date:
            normalized_bill_date = normalized_bill_date.split(' ')[0]
        
        print(f"정규화된 날짜: '{normalized_bill_date}'")
        
        rows = []
        for _, row in df.iterrows():
            # DataFrame의 날짜 형식 통일 (시간 부분 제거)
            row_bill_date = str(row['계산서 발행일'])
            if ' ' in row_bill_date:
                row_bill_date = row_bill_date.split(' ')[0]
            
            # 디버깅을 위한 로그
            row_customer = str(row['고객사명']).strip()
            row_client = str(row['거래처명']).strip()
            
            print(f"비교: '{row_bill_date}' == '{normalized_bill_date}', '{row_customer}' == '{customer_name}', '{row_client}' == '{client_name}'")
            
            # 문자열 비교를 위해 str로 변환하고 공백 제거
            if (row_bill_date == normalized_bill_date and 
                row_customer == customer_name.strip() and 
                row_client == client_name.strip()):
                rows.append({col: row.get(col, '') for col in self.purchase_cols})
                print(f"매칭되는 행 발견: {row_bill_date}, {row_customer}, {row_client}")

        print(f"필터링된 행 수: {len(rows)}")
        if not rows:
            print(f"매칭되는 데이터가 없음: 계산서 발행일={bill_date}, 고객사명={customer_name}, 거래처명={client_name}")
            tk.messagebox.showinfo('알림', f'계산서 발행일: {bill_date}, 고객사명: {customer_name}, 거래처명: {client_name}인 데이터가 없습니다.')
            return
        try:
            from modules.quote.quote_simple import QuoteSimpleDialog
            # MainApp의 username을 전달
            username = getattr(self, 'username', '테스트유저')
            
            # 견적서 탭의 refresh_quote_data 함수에 접근하기 위한 콜백 함수 정의
            def combined_callback():
                # 고객사 관리 데이터 새로고침
                self.load_data()
                # 견적서 탭 데이터 새로고침
                if self.quote_gui:
                    try:
                        self.quote_gui.refresh_quote_data()
                    except Exception as e:
                        print(f"견적서 탭 새로고침 실패: {e}")
            
            print(f"견적 작성(기존): 다이얼로그 시작 - 고객사: {customer_name}, 거래처: {client_name}")
            QuoteSimpleDialog(self.frame, past_rows=rows, username=username, callback=combined_callback)
            print("견적 작성(기존): 다이얼로그 생성 완료")
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f"견적 작성(기존): 다이얼로그 생성 실패 - {e}")
            tk.messagebox.showerror('오류', f'견적 작성(기존) 창을 여는 데 실패했습니다:\n{e}')

    def on_category_search(self):
        """대분류 콤보박스에서 엔터키로 검색"""
        entered_category = self.category_cb.get().strip()
        if entered_category:
            # 디버깅용 로그 추가
            print(f"대분류 검색: '{entered_category}', 전체 대분류 수: {len(self.all_categories)}")
            
            if not self.all_categories:
                print("경고: 대분류 목록이 비어있습니다!")
                messagebox.showwarning("검색 오류", "대분류 목록이 비어있습니다.")
                return
                
            # 대분류로 검색 (부분 일치)
            matching_categories = [c for c in self.all_categories if entered_category.lower() in c.lower()]
            print(f"매칭되는 대분류: {matching_categories}")
            
            if matching_categories:
                # 검색 결과를 콤보박스에 설정
                self.category_cb['values'] = ['전체'] + matching_categories
                # 첫 번째 매칭되는 대분류를 설정
                self.category_cb.set(matching_categories[0])
                # 드롭다운 리스트 표시
                self.category_cb.event_generate('<Down>')
                self.category_cb.event_generate('<Up>')
                
                # 검색 결과가 있으면 자동으로 검색 실행
                self.on_search()
                
                print(f"대분류 검색 완료: '{matching_categories[0]}' 선택됨")
                
                # 검색 결과가 여러 개일 때 안내 메시지
                if len(matching_categories) > 1:
                    messagebox.showinfo("검색 결과", f"'{entered_category}'와 일치하는 대분류 {len(matching_categories)}개를 찾았습니다.\n첫 번째 결과가 선택되었습니다.")
            else:
                # 매칭되는 대분류가 없으면 메시지 표시
                messagebox.showinfo("검색 결과", f"'{entered_category}'와 일치하는 대분류를 찾을 수 없습니다.\n\n다른 키워드로 검색해보세요.")
                # 입력값을 그대로 유지하되 검색은 실행하지 않음
                self.category_cb.set(entered_category)
