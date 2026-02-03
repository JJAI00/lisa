# purchase_sales_info.py

import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import pandas as pd
import calendar
from datetime import datetime
from modules.renewal import renewal_search as rs
from .purchase_sales_modify import PurchaseSalesModifyDialog
from integrations.lisa_logging import setup_logger

# 모듈 로거 설정
logger = setup_logger('purchase_sales_info')

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

class PurchaseSalesInfo:
    """매입/매출 탭 내 임베디드 UI - renewal_list 전체 컬럼 표시 및 기간 필터"""
    def __init__(self, parent_frame, username=None, data_loader=None):
        self.parent_frame = parent_frame
        self.username = username or '테스트유저'
        self.data_loader = data_loader
        
        # 데이터 로더가 있으면 사용, 없으면 기존 방식 사용
        if self.data_loader:
            try:
                self.renewals_df = self.data_loader.get_renewal_data()
                
                # 데이터가 비어있으면 기존 방식으로 로드
                if self.renewals_df.empty:
                    print("매입/매출: 데이터 로더에서 데이터를 가져올 수 없어 기존 방식으로 로드합니다.")
                    self._load_data_directly()
                else:
                    print(f"매입/매출: 데이터 로더에서 데이터를 성공적으로 가져왔습니다: {len(self.renewals_df)}행")
            except Exception as e:
                print(f"매입/매출: 데이터 로더 오류 - {e}, 기존 방식으로 로드합니다.")
                self._load_data_directly()
        else:
            # 기존 방식 (하위 호환성)
            print("매입/매출: 기존 방식으로 데이터를 로드합니다.")
            self._load_data_directly()
        
        # 계산서 발행일 컬럼을 날짜로 변환 (원본은 그대로 둠)
        if '계산서 발행일' in self.renewals_df.columns:
            self.renewals_df['계산서_날짜'] = pd.to_datetime(
                self.renewals_df['계산서 발행일'], format='%Y-%m-%d', errors='coerce'
            )

        # 필터용 값들
        self.all_customers = sorted(
            self.renewals_df['고객사명'].dropna().unique().tolist()
        ) if '고객사명' in self.renewals_df.columns else []
        self.all_clients   = sorted(
            self.renewals_df['거래처명'].dropna().unique().tolist()
        ) if '거래처명' in self.renewals_df.columns else []
        self.all_categories = sorted(
            self.renewals_df['대분류'].dropna().unique().tolist()
        ) if '대분류' in self.renewals_df.columns else []
        self.all_sales     = sorted(
            self.renewals_df['영업사원'].dropna().unique().tolist()
        ) if '영업사원' in self.renewals_df.columns else []
        self.numeric_cols  = self.renewals_df.select_dtypes(
            include=['number']
        ).columns.tolist()

        # 정렬/기간 상태 초기화
        self.sort_column  = None
        self.sort_reverse = False
        self.last_start   = None
        self.last_end     = None

        # 프레임 생성 및 위젯 배치
        self.frame = tk.Frame(self.parent_frame)
        self.frame.pack(fill='both', expand=True)

        self.setup_widgets()
        self.load_filters()
        # 앱 실행 시 기본 "이번달" 표시
        self.load_period('this_month')
        
        # 데이터 로딩 상태 출력
        print(f"매입/매출 탭 초기화 완료")
        print(f"renewals_df 크기: {len(self.renewals_df)}행")
    
    def on_full_data_loaded(self):
        """전체 데이터 로딩 완료 시 호출되는 메서드"""
        if self.data_loader:
            try:
                # 데이터 로더에서 최신 데이터 가져오기
                self.renewals_df = self.data_loader.get_renewal_data()
                
                # 데이터가 비어있으면 기존 방식으로 로드
                if self.renewals_df.empty:
                    print("매입/매출: 전체 데이터 로딩에서 데이터를 가져올 수 없어 기존 방식으로 로드합니다.")
                    self._load_data_directly()
                
                # 필터 옵션 업데이트
                self.load_filters()
                
                print(f"매입/매출 탭 전체 데이터 로딩 완료")
                print(f"renewals_df 크기: {len(self.renewals_df)}행")
            except Exception as e:
                print(f"매입/매출: 전체 데이터 로딩 오류 - {e}")
                # 에러 발생 시 기존 데이터 유지

    def setup_widgets(self):
        # 상단 헤더
        header_frame = tk.Frame(self.frame, bg="#F7F9FB", height=60)
        header_frame.pack(fill='x', pady=(0, 20))
        header_frame.pack_propagate(False)
        
        # 헤더 제목
        title_label = tk.Label(header_frame, text="매입/매출 관리", 
                              fg="#1F2937", bg="#F7F9FB")
        title_label.pack(side='left', padx=20, pady=15)
        
        # 사용자 정보 (우측)
        user_label = tk.Label(header_frame, text=f"👤 {self.username}", 
                             fg="#6B7280", bg="#F7F9FB")
        user_label.pack(side='right', padx=20, pady=15)

        # 검색/필터 카드 섹션
        filter_section = tk.Frame(self.frame, bg="#F7F9FB")
        filter_section.pack(fill='x', padx=20, pady=(0, 20))

        # 첫 번째 카드: 날짜 선택 및 기간
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
        
        # 날짜 선택
        date_frame = tk.Frame(card1_content, bg="white")
        date_frame.pack(fill='x', pady=(0, 10))
        tk.Label(date_frame, text="조회 기준일:", 
                fg="#374151", bg="white").pack(side='left')
        self.date_entry = DateEntry(
            date_frame,
            width=12,
            date_pattern='yyyy-mm-dd'
        )
        self.date_entry.set_date(datetime.today())
        self.date_entry.pack(side='left', padx=5)
        # 날짜 선택 시 "이번달" 자동 조회
        self.date_entry.bind(
            "<<DateEntrySelected>>",
            lambda e: self.load_period('this_month')
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
                                 command=lambda: self.load_period('past_year'),
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
        
        # 고객사명, 거래처명 필터를 드롭다운+입력+엔터로 검색되도록 개선
        tk.Label(row1, text="고객사명:", font=("맑은 고딕", 10), fg="#374151", bg="white").pack(side='left')
        self.customer_cb = ttk.Combobox(
            row1,
            state='normal',
            width=20,
            postcommand=self.update_customer_list
        )
        self.customer_cb.pack(side='left', padx=5)
        self.customer_cb.bind('<<ComboboxSelected>>', lambda e: self.load_data())
        self.customer_cb.bind('<Return>', lambda e: self.on_customer_search())

        tk.Label(row1, text="거래처명:", font=("맑은 고딕", 10), fg="#374151", bg="white").pack(side='left', padx=(20,0))
        self.client_cb = ttk.Combobox(
            row1,
            state='normal',
            width=20,
            postcommand=self.update_client_list
        )
        self.client_cb.pack(side='left', padx=5)
        self.client_cb.bind('<<ComboboxSelected>>', lambda e: self.load_data())
        self.client_cb.bind('<Return>', lambda e: self.on_client_search())

        # 두 번째 행: 대분류, 영업사원
        row2 = tk.Frame(card2_content, bg="white")
        row2.pack(fill='x', pady=(0, 10))
        
        tk.Label(row2, text="대분류:", font=("맑은 고딕", 10), fg="#374151", bg="white").pack(side='left')
        self.category_cb = ttk.Combobox(
            row2,
            state='normal',
            width=20,
            postcommand=self.update_category_list
        )
        self.category_cb.pack(side='left', padx=5)
        self.category_cb.bind('<<ComboboxSelected>>', lambda e: self.load_data())
        self.category_cb.bind('<Return>', lambda e: self.on_category_search())

        tk.Label(row2, text="영업사원:", font=("맑은 고딕", 10), fg="#374151", bg="white").pack(side='left', padx=(20,0))
        self.sales_cb = ttk.Combobox(
            row2,
            state='normal',
            width=20,
            postcommand=self.update_sales_list
        )
        self.sales_cb.pack(side='left', padx=5)
        self.sales_cb.bind('<<ComboboxSelected>>', lambda e: self.load_data())
        self.sales_cb.bind('<Return>', lambda e: self.on_sales_search())

        # 세 번째 행: 검색 버튼만
        row3 = tk.Frame(card2_content, bg="white")
        row3.pack(fill='x')
        
        search_btn = tk.Button(row3, text="🔍 검색", command=self.load_data,
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
        
        # 계산서/결제/매입 수정 버튼
        modify_btn = tk.Button(action_buttons, text="💰 계산서/결제/매입",
                              command=self.open_modify_dialog,
                              font=("맑은 고딕", 10, "bold"), bg="#F59E0B", fg="white",
                              relief="flat", padx=15, pady=4)
        modify_btn.pack(fill='x', pady=(0, 5))
        modify_btn.bind('<Enter>', lambda e: modify_btn.configure(bg="#D97706"))
        modify_btn.bind('<Leave>', lambda e: modify_btn.configure(bg="#F59E0B"))
        
        # 만료일 수정 버튼 (결제/매입 단가 수정 밑에 배치)
        expdate_modify_btn = tk.Button(action_buttons, text="📅 만료일 수정",
                                      command=self.open_expdate_modify_dialog,
                                      font=("맑은 고딕", 10, "bold"), bg="#8B5CF6", fg="white",
                                      relief="flat", padx=15, pady=4)
        expdate_modify_btn.pack(fill='x', pady=(0, 5))
        expdate_modify_btn.bind('<Enter>', lambda e: expdate_modify_btn.configure(bg="#7C3AED"))
        expdate_modify_btn.bind('<Leave>', lambda e: expdate_modify_btn.configure(bg="#8B5CF6"))

        # 데이터 테이블 섹션
        table_section = tk.Frame(self.frame, bg="#F7F9FB")
        table_section.pack(fill='both', expand=True, padx=20, pady=(0, 20))
        
        # 테이블 헤더
        table_header = tk.Frame(table_section, bg="white", relief="solid", bd=1)
        table_header.pack(fill='x', pady=(0, 1))
        tk.Label(table_header, text="📊 매입/매출 데이터", font=("맑은 고딕", 14, "bold"), 
                fg="#1F2937", bg="white").pack(side='left', padx=15, pady=10)
        
        # 테이블 컨테이너
        tree_container = tk.Frame(table_section, bg="white", relief="solid", bd=1)
        tree_container.pack(fill='both', expand=True)

        # 컬럼 순서 조정: 계산서 발행일을 제일 앞으로, 그 다음에 영업사원
        all_cols = list(self.renewals_df.columns)
        if '계산서 발행일' in all_cols and '영업사원' in all_cols:
            # 계산서 발행일과 영업사원을 제외한 나머지 컬럼들
            other_cols = [col for col in all_cols if col not in ['계산서 발행일', '영업사원']]
            # 원하는 순서로 재배열
            self.cols = ['계산서 발행일', '영업사원'] + other_cols
        else:
            self.cols = all_cols
            
        self.tree = ttk.Treeview(
            tree_container, columns=self.cols, show='headings',
            selectmode='browse'
        )
        vsb = ttk.Scrollbar(
            tree_container, orient='vertical', command=self.tree.yview
        )
        hsb = ttk.Scrollbar(
            tree_container, orient='horizontal', command=self.tree.xview
        )
        self.tree.configure(
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set
        )
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        self.tree.pack(fill='both', expand=True, padx=1, pady=1)

        # 컬럼 크기 조절 기능 추가
        from ui.resizable_treeview import ResizableTreeview
        self.resizable_tree = ResizableTreeview(self.tree, "purchase_sales_table")

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

        for col in self.cols:
            anchor = 'e' if col in self.numeric_cols else 'w'
            self.tree.heading(
                col, text=col,
                command=lambda c=col: self.sort_purchase(c)
            )
            
            # 컬럼별 너비 설정
            if col == '계산서 발행일':
                width = 120
            elif col == '영업사원':
                width = 100
            elif col in ['만료일']:
                width = 120
            else:
                width = 100
                
            self.tree.column(
                col, width=width,
                anchor=anchor, stretch=False
            )

        # 금액(숫자) 컬럼 리스트 추가
        self.cost_cols = [
            '판매 단가','판매 합계','원가','원가계',
            '실 매입 단가','실 매입 합계','이익액','실 이익액',
            '입금액','잔액'
        ]

    def load_filters(self):
        """콤보박스 초기 옵션 로드"""
        self.customer_cb['values'] = ['전체'] + self.all_customers
        self.customer_cb.set('전체')
        self.client_cb['values']   = ['전체'] + self.all_clients
        self.client_cb.set('전체')
        self.category_cb['values'] = ['전체'] + self.all_categories
        self.category_cb.set('전체')
        self.sales_cb['values']    = ['전체'] + self.all_sales
        self.sales_cb.set('전체')

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
                
            # 고객사명으로 검색
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
                
            # 거래처명으로 검색
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

    def on_category_search(self):
        """대분류 콤보박스에서 엔터키로 검색"""
        entered_category = self.category_cb.get().strip()
        if entered_category:
            # 디버깅용 로그 추가
            print(f"대분류 검색: '{entered_category}', 전체 대분류 수: {len(self.all_categories)}")
            
            if not self.all_categories:
                print("경고: 대분류 목록이 비어있습니다!")
                return
                
            # 대분류로 검색
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
            else:
                # 매칭되는 대분류가 없으면 입력값 그대로 사용
                self.category_cb.set(entered_category)

    def on_sales_search(self):
        """영업사원 콤보박스에서 엔터키로 검색"""
        entered_sales = self.sales_cb.get().strip()
        if entered_sales:
            # 디버깅용 로그 추가
            print(f"영업사원 검색: '{entered_sales}', 전체 영업사원 수: {len(self.all_sales)}")
            
            if not self.all_sales:
                print("경고: 영업사원 목록이 비어있습니다!")
                return
                
            # 영업사원으로 검색
            matching_sales = [c for c in self.all_sales if entered_sales.lower() in c.lower()]
            print(f"매칭되는 영업사원: {matching_sales}")
            
            if matching_sales:
                # 검색 결과를 콤보박스에 설정
                self.sales_cb['values'] = ['전체'] + matching_sales
                # 첫 번째 매칭되는 영업사원을 설정
                self.sales_cb.set(matching_sales[0])
                # 드롭다운 리스트 표시
                self.sales_cb.event_generate('<Down>')
                self.sales_cb.event_generate('<Up>')
            else:
                # 매칭되는 영업사원이 없으면 입력값 그대로 사용
                self.sales_cb.set(entered_sales)

    def refresh_data(self):
        """데이터 새로고침"""
        print("매입/매출: 데이터 새로고침 중...")
        if self.data_loader:
            # 데이터 로더에서 최신 데이터 가져오기
            self.renewals_df = self.data_loader.get_renewal_data()
            print(f"매입/매출: 데이터 로더에서 데이터 새로고침 완료 - {len(self.renewals_df)}행")
        else:
            # 기존 방식으로 데이터 새로고침
            self._load_data_directly()
            print("매입/매출: 기존 방식으로 데이터 새로고침 완료")
        
        # 필터 목록 업데이트
        self.load_filters()
        # 현재 필터 적용하여 데이터 로드
        self.load_data()

    def load_all_data(self):
        """전체 데이터 로드 (기간 제한 없음)"""
        print("매입/매출: 전체 데이터 로드 버튼 클릭")
        try:
            # 필터 필드 초기화
            self.customer_cb.set('전체')
            self.client_cb.set('전체')
            self.category_cb.set('전체')
            self.sales_cb.set('전체')
            
            # 전체 데이터를 트리뷰에 표시 (기간 필터링 없이)
            if self.renewals_df is not None and not self.renewals_df.empty:
                print(f"매입/매출: 전체 데이터 로드: {len(self.renewals_df)}행")
                
                # 기간 필터링 없이 전체 데이터 표시
                df = self.renewals_df.copy()
                
                # 트리뷰에 데이터 표시 (컬럼 순서에 맞춰서)
                self.tree.delete(*self.tree.get_children())
                for _, row in df.iterrows():
                    vals = []
                    for col in self.cols:
                        v = row.get(col, '')
                        # 만료일, 계산서 발행일은 날짜 포맷팅
                        if col in ['만료일', '계산서 발행일']:
                            if pd.notna(v):
                                if isinstance(v, str):
                                    # 시간 부분이 포함된 경우 날짜만 추출
                                    if ' ' in v:
                                        v = v.split(' ')[0]
                                elif hasattr(v, 'strftime'):
                                    # NaT 값 체크
                                    try:
                                        v = v.strftime('%Y-%m-%d')
                                    except:
                                        v = ''
                        vals.append(str(v) if v is not None else '')
                    self.tree.insert('', 'end', values=vals)
                
                messagebox.showinfo("완료", f"전체 데이터 {len(self.renewals_df)}행을 로드했습니다.")
            else:
                print("매입/매출: 로드할 데이터가 없습니다.")
                messagebox.showwarning("경고", "로드할 데이터가 없습니다.")
        except Exception as e:
            print(f"매입/매출: 전체 데이터 로드 오류: {e}")
            messagebox.showerror("오류", f"전체 데이터 로드 중 오류가 발생했습니다:\n{e}")

    def load_period(self, mode):
        """date_entry 기준으로 기간 설정 후 load_data 호출"""
        base = self.date_entry.get_date()
        if mode == 'today':
            start = end = base
        elif mode == 'this_month':
            start = base.replace(day=1)
            end   = base.replace(day=calendar.monthrange(base.year, base.month)[1])
        elif mode == 'next_month':
            year  = base.year + (base.month // 12)
            month = (base.month % 12) + 1
            start = base.replace(year=year, month=month, day=1)
            end   = start.replace(day=calendar.monthrange(year, month)[1])
        elif mode == 'one_year':
            result = rs.period_one_year(base)
            if result and isinstance(result, tuple) and len(result) == 2:
                start, end = result
            else:
                start = base
                end = base
        elif mode == 'past_year':
            result = rs.period_past_year(base)
            if result and isinstance(result, tuple) and len(result) == 2:
                start, end = result
            else:
                start = base
                end = base
        else:
            start = end = base
        # 항상 pd.to_datetime()으로 변환해서 저장
        self.last_start = pd.to_datetime(start)
        self.last_end = pd.to_datetime(end)
        self.load_data()

    def load_data(self):
        """last_start~last_end 기준 데이터 필터링 및 Treeview 갱신"""
        df = self.renewals_df.copy()
        # 만료일, 계산서 발행일을 날짜로 변환 (공백/NaN/None은 NaT)
        if '만료일' in df.columns:
            df['만료일'] = pd.to_datetime(df['만료일'], errors='coerce')
        if '계산서 발행일' in df.columns:
            df['계산서 발행일'] = pd.to_datetime(df['계산서 발행일'], format='%Y-%m-%d', errors='coerce')
        # 기간 필터링: 계산서 발행일 기준, 비교값을 pd.to_datetime()으로 변환
        start_dt = self.last_start
        end_dt = self.last_end
        if start_dt is not None and end_dt is not None and '계산서 발행일' in df.columns:
            mask = (df['계산서 발행일'] >= start_dt) & (df['계산서 발행일'] <= end_dt)
            df = df[mask]
        
        # 필터 적용
        cust   = self.customer_cb.get().strip()
        client = self.client_cb.get().strip()
        cat    = self.category_cb.get().strip()
        sales  = self.sales_cb.get().strip()
        
        if cust and cust != '전체':
            df = df[df['고객사명'].str.contains(cust, na=False)]
        if client and client != '전체':
            df = df[df['거래처명'].str.contains(client, na=False)]
        if cat and cat != '전체' and '대분류' in df.columns:
            df = df[df['대분류'].str.contains(cat, na=False)]
        if sales and sales != '전체' and '영업사원' in df.columns:
            df = df[df['영업사원'].str.contains(sales, na=False)]
            
        self.tree.delete(*self.tree.get_children())
        for _, row in df.iterrows():
            vals = []
            for col in self.cols:
                v = row.get(col, '')
                # 만료일, 계산서 발행일은 날짜 포맷팅
                if col in ['만료일', '계산서 발행일']:
                    if pd.notna(v):
                        if isinstance(v, str):
                            # 시간 부분이 포함된 경우 날짜만 추출
                            if ' ' in v:
                                v = v.split(' ')[0]
                            else:
                                v = v
                        elif hasattr(v, 'strftime'):
                            v = v.strftime('%Y-%m-%d')
                    else:
                        v = ''
                # 금액 컬럼은 천 단위 콤마
                elif col in self.cost_cols:
                    try:
                        v = f"{safe_int_convert(v):,}"
                    except:
                        pass
                vals.append(v)
            self.tree.insert('', 'end', values=vals)

    def sort_purchase(self, col):
        """컬럼 클릭 시 정렬"""
        if self.sort_column == col:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column  = col
            self.sort_reverse = False

        self.renewals_df = self.renewals_df.sort_values(
            by=col,
            ascending=not self.sort_reverse,
            key=lambda x: x if col in self.numeric_cols else x.fillna('').astype(str)
        )
        self.load_data()
        
    def open_modify_dialog(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("선택 필요", "수정할 행을 먼저 선택하세요.")
            return
            
        item = self.tree.item(selected[0])
        row_values = item['values']
        
        # 컬럼명과 값 매핑
        row_data = {}
        for i, col in enumerate(self.cols):
            if i < len(row_values):
                val = row_values[i]
                # 빈 값이나 None 처리
                if val is None or val == '' or str(val).strip() == '':
                    val = ''
                row_data[col] = val
            else:
                row_data[col] = ''
        

        
        PurchaseSalesModifyDialog(self.frame, row_data, refresh_callback=self.load_data)

    def open_expdate_modify_dialog(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("선택 필요", "수정할 행을 먼저 선택하세요.")
            return
            
        item = self.tree.item(selected[0])
        row_values = item['values']
        
        # 컬럼명과 값 매핑
        row_data = {}
        for i, col in enumerate(self.cols):
            if i < len(row_values):
                val = row_values[i]
                # 빈 값이나 None 처리
                if val is None or val == '' or str(val).strip() == '':
                    val = ''
                row_data[col] = val
            else:
                row_data[col] = ''
        
        from purchase_sales_expdate_modify import PurchaseSalesExpdateModifyDialog
        PurchaseSalesExpdateModifyDialog(self.frame, row_data, refresh_callback=self.load_data)

    def _load_data_directly(self):
        """기존 방식으로 데이터 직접 로드"""
        try:
            creds = rs._authorize()
            # Google Sheet에서 리뉴얼 목록 로드
            sh = creds.open_by_key(rs.RENEWAL_LIST_SHEET_ID)
            ws = sh.worksheet(rs.RENEWAL_SHEET_NAME)
            self.renewals_df = pd.DataFrame(ws.get_all_records())
            self.renewals_df.columns = self.renewals_df.columns.str.strip()
            print(f"매입/매출: 직접 로드 완료 - {len(self.renewals_df)}행")
        except Exception as e:
            print(f"매입/매출: 직접 로드 실패 - {e}")
            self.renewals_df = pd.DataFrame()



