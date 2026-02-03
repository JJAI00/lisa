import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime, timedelta
import gspread
from google.oauth2.service_account import Credentials
import matplotlib
matplotlib.use('TkAgg')  # 백엔드 명시적 설정
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
import warnings
warnings.filterwarnings('ignore')

class DetailWindow:
    """상세 정보 창"""
    def __init__(self, parent, title, data, foundry_df):
        self.window = tk.Toplevel(parent)
        self.window.title(title)
        self.window.geometry("1200x600")
        self.window.configure(bg='white')
        
        # 창을 부모 창의 중앙에 위치
        self.window.transient(parent)
        self.window.grab_set()
        
        # 데이터
        self.data = data
        self.foundry_df = foundry_df
        
        self.setup_ui()
    
    def setup_ui(self):
        """UI 설정"""
        # 헤더
        header_frame = tk.Frame(self.window, bg='white')
        header_frame.pack(fill='x', padx=10, pady=10)
        
        title_label = tk.Label(header_frame, text="상세 정보", 
                              font=("맑은 고딕", 16, "bold"), bg='white', fg='#2c3e50')
        title_label.pack(side='left')
        
        # 테이블 프레임
        table_frame = tk.Frame(self.window, bg='white')
        table_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 테이블 생성
        columns = ['FY', 'Quarter', 'Partner Name', 'Account', 'Product', 'Product Type 1', 
                  'Q\'ty', 'Unit Price', 'Revenue', 'Status', '만료일']
        
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=20)
        
        # 컬럼 설정
        for col in columns:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_detail_treeview(c, False))
            if col in ['Q\'ty', 'Unit Price', 'Revenue']:
                self.tree.column(col, width=100, anchor='e')
            elif col in ['FY', 'Quarter']:
                self.tree.column(col, width=80, anchor='center')
            else:
                self.tree.column(col, width=120, anchor='w')
        
        # 스크롤바
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        
        # 초기 데이터 로드
        self.load_data()
    
    def sort_detail_treeview(self, col, reverse):
        """DetailWindow 트리뷰 정렬 함수"""
        l = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        
        # 숫자 정렬을 위한 함수
        def try_float(s):
            try:
                return float(s.replace(',', ''))
            except:
                return 0
        
        # 숫자 컬럼인지 확인
        if col in ['Q\'ty', 'Unit Price', 'Revenue']:
            l.sort(key=lambda t: try_float(t[0]), reverse=reverse)
        else:
            l.sort(reverse=reverse)
        
        # 재정렬
        for index, (val, k) in enumerate(l):
            self.tree.move(k, '', index)
        
        # 다음 클릭을 위해 반전
        self.tree.heading(col, command=lambda: self.sort_detail_treeview(col, not reverse))
    
    def load_data(self):
        """데이터 로드"""
        # 기존 데이터 삭제
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 데이터 삽입
        for _, row in self.data.iterrows():
            values = []
            for col in ['FY', 'Quarter', 'Partner Name', 'Account', 'Product', 'Product Type 1', 
                       'Q\'ty', 'Unit Price', 'Revenue', 'Status', '만료일']:
                if col in row:
                    value = row[col]
                    if pd.notna(value):
                        if col in ['Q\'ty', 'Unit Price', 'Revenue']:
                            try:
                                values.append(f"{float(value):,.0f}")
                            except:
                                values.append(str(value))
                        else:
                            values.append(str(value))
                    else:
                        values.append('')
                else:
                    values.append('')
            
            self.tree.insert('', 'end', values=values)

class FoundryDashboard:
    def __init__(self, parent_frame):
        self.parent_frame = parent_frame
        self.foundry_df = pd.DataFrame()
        self.foundry_rev_df = pd.DataFrame()
        self.current_tab = "전체 실적"
        self.selected_years = ["전체"]
        self.selected_quarters = ["전체"]
        
        # 무한 루프 방지를 위한 플래그들
        self._updating_year = False
        self._updating_quarter = False
        self._updating_dashboard = False
        self.update_timer = None
        
        # UI 컴포넌트
        self.tab_control = None
        self.year_combo = None
        self.quarter_combo = None
        self.summary_frame = None
        self.chart_frame = None
        self.table_frame = None
        
        # 구글 시트 연결
        self.connect_google_sheets()
        
        # 데이터 로드
        self.load_data()
        
        # UI 설정
        self.setup_ui()
    
    def connect_google_sheets(self):
        """구글 시트 연결"""
        try:
            # 구글 시트 API 설정
            SCOPES = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
            
            # 서비스 계정 키 파일 경로
            KEY_FILE = 'google_sheet_key/renewal-bot-463605-b8f41de8fbdb.json'
            
            # 인증
            creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
            self.gc = gspread.authorize(creds)
            
            # 시트 열기
            self.spreadsheet = self.gc.open_by_key('1C5p4HBn9G8MpPdQf6Q8FqaM4lgUzzfftBqG-qQhLa7Y')
            
            print("구글 시트 연결 성공")
            
        except Exception as e:
            print(f"구글 시트 연결 실패: {e}")
            messagebox.showerror("오류", f"구글 시트 연결에 실패했습니다:\n{e}")
    
    def load_data(self):
        """데이터 로드"""
        try:
            # Foundry 시트 (실적 데이터)
            foundry_ws = self.spreadsheet.worksheet('Foundry')
            foundry_data = foundry_ws.get_all_records()
            self.foundry_df = pd.DataFrame(foundry_data)
            
            # Foundry_Rev 시트 (목표 데이터)
            foundry_rev_ws = self.spreadsheet.worksheet('Foundry_Rev')
            foundry_rev_data = foundry_rev_ws.get_all_records()
            self.foundry_rev_df = pd.DataFrame(foundry_rev_data)
            
            # 실제 컬럼명 출력
            print(f"Foundry 시트 컬럼명: {list(self.foundry_df.columns)}")
            print(f"Foundry_Rev 시트 컬럼명: {list(self.foundry_rev_df.columns)}")
            
            # 데이터 전처리
            self.preprocess_data()
            
            print(f"데이터 로드 완료 - Foundry: {len(self.foundry_df)}행, Foundry_Rev: {len(self.foundry_rev_df)}행")
            
        except Exception as e:
            print(f"데이터 로드 실패: {e}")
            messagebox.showerror("오류", f"데이터 로드에 실패했습니다:\n{e}")
    
    def preprocess_data(self):
        """데이터 전처리"""
        try:
            # Foundry 데이터 전처리 (실적)
            if not self.foundry_df.empty:
                # Unit Price를 숫자로 변환
                if 'Unit Price' in self.foundry_df.columns:
                    self.foundry_df['Unit Price'] = pd.to_numeric(self.foundry_df['Unit Price'], errors='coerce')
                
                # Q'ty를 숫자로 변환
                if "Q'ty" in self.foundry_df.columns:
                    self.foundry_df["Q'ty"] = pd.to_numeric(self.foundry_df["Q'ty"], errors='coerce')
                
                # FY와 Quarter를 문자열로 변환
                if 'FY' in self.foundry_df.columns:
                    self.foundry_df['FY'] = self.foundry_df['FY'].astype(str)
                
                if 'Quarter' in self.foundry_df.columns:
                    self.foundry_df['Quarter'] = self.foundry_df['Quarter'].astype(str)
            
                # Product Type 1 컬럼이 없으면 빈 문자열로 추가
                if 'Product Type 1' not in self.foundry_df.columns:
                    self.foundry_df['Product Type 1'] = ''
                
                # Product 컬럼이 없으면 빈 문자열로 추가
                if 'Product' not in self.foundry_df.columns:
                    self.foundry_df['Product'] = ''
                
                # 매출 계산 (Unit Price가 이미 단가×수량으로 계산된 값)
                if 'Unit Price' in self.foundry_df.columns:
                    self.foundry_df['Revenue'] = self.foundry_df['Unit Price']
                else:
                    self.foundry_df['Revenue'] = 0
            
            # Foundry_Rev 데이터 전처리 (목표)
            if not self.foundry_rev_df.empty:
                # 연도를 문자열로 변환
                if '연도' in self.foundry_rev_df.columns:
                    self.foundry_rev_df['연도'] = self.foundry_rev_df['연도'].astype(str)
                
                # 구분을 문자열로 변환
                if '구분' in self.foundry_rev_df.columns:
                    self.foundry_rev_df['구분'] = self.foundry_rev_df['구분'].astype(str)
                
                # Q1, Q2, Q3, Q4, TTL을 숫자로 변환
                for col in ['Q1', 'Q2', 'Q3', 'Q4', 'TTL']:
                    if col in self.foundry_rev_df.columns:
                        self.foundry_rev_df[col] = pd.to_numeric(self.foundry_rev_df[col], errors='coerce')
            
                # 중복된 연도 데이터 처리 (첫 번째 값만 유지)
                self.foundry_rev_df = self.foundry_rev_df.drop_duplicates(subset=['연도', '구분'], keep='first')
                
                print(f"Foundry_Rev 데이터 샘플:")
                print(self.foundry_rev_df.head())
                
                # 전체 목표 합계 확인
                total_target = self.foundry_rev_df['TTL'].sum()
                print(f"Foundry_Rev 전체 목표 합계: {total_target:,.2f}")
            
            print("데이터 전처리 완료")
            
        except Exception as e:
            print(f"데이터 전처리 실패: {e}")
            import traceback
            traceback.print_exc()
    
    def setup_ui(self):
        """UI 설정"""
        # 메인 컨테이너
        main_container = tk.Frame(self.parent_frame, bg='white')
        main_container.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 헤더 섹션
        self.setup_header(main_container)
        
        # 요약 지표 섹션
        self.setup_summary_metrics(main_container)
        
        # 탭 컨트롤
        self.setup_tabs(main_container)
        
        # 초기 데이터 로드
        self.update_dashboard()
    
    def setup_header(self, parent):
        """헤더 설정"""
        header_frame = tk.Frame(parent, bg='white')
        header_frame.pack(fill='x', pady=(0, 20))
        
        # 제목
        title_label = tk.Label(header_frame, text="Foundry 대시보드", 
                              font=("맑은 고딕", 20, "bold"), bg='white', fg='#2c3e50')
        title_label.pack(side='left')
        
        # 필터 프레임
        filter_frame = tk.Frame(header_frame, bg='white')
        filter_frame.pack(side='right')
        
        # 필터 모드 선택
        mode_frame = tk.Frame(filter_frame, bg='white')
        mode_frame.pack(side='left', padx=(0, 20))
        
        tk.Label(mode_frame, text="필터 모드:", font=("맑은 고딕", 10), bg='white').pack(side='left', padx=(0, 5))
        
        self.filter_mode = tk.StringVar(value="기본")
        tk.Radiobutton(mode_frame, text="기본", variable=self.filter_mode, value="기본", 
                      command=self.on_filter_mode_changed, bg='white').pack(side='left', padx=(0, 5))
        tk.Radiobutton(mode_frame, text="비교", variable=self.filter_mode, value="비교", 
                      command=self.on_filter_mode_changed, bg='white').pack(side='left')
        
        # 기본 필터 프레임
        self.basic_filter_frame = tk.Frame(filter_frame, bg='white')
        self.basic_filter_frame.pack(side='left', padx=(0, 20))
        
        # 연도 필터 (다중 선택)
        year_filter_frame = tk.Frame(self.basic_filter_frame, bg='white')
        year_filter_frame.pack(side='left', padx=(0, 20))
        
        tk.Label(year_filter_frame, text="연도:", font=("맑은 고딕", 12), bg='white').pack(side='left', padx=(0, 5))
        
        # 연도 옵션 생성
        if not self.foundry_df.empty:
            years = ["전체"] + sorted(self.foundry_df['FY'].unique().tolist())
        elif not self.foundry_rev_df.empty:
            years = ["전체"] + sorted(self.foundry_rev_df['연도'].unique().tolist())
        else:
            years = ["전체"]
        
        # 다중 선택을 위한 Listbox로 변경
        year_listbox_frame = tk.Frame(year_filter_frame, bg='white')
        year_listbox_frame.pack(side='left')
        
        self.year_listbox = tk.Listbox(year_listbox_frame, selectmode='multiple', height=4, width=12)
        self.year_listbox.pack(side='left')
        
        # 연도 옵션 추가
        for year in years:
            self.year_listbox.insert(tk.END, year)
        
        # 스크롤바
        year_scrollbar = tk.Scrollbar(year_listbox_frame, orient="vertical", command=self.year_listbox.yview)
        self.year_listbox.configure(yscrollcommand=year_scrollbar.set)
        year_scrollbar.pack(side='right', fill='y')
        
        # 분기 필터 (다중 선택)
        quarter_filter_frame = tk.Frame(self.basic_filter_frame, bg='white')
        quarter_filter_frame.pack(side='left')
        
        tk.Label(quarter_filter_frame, text="분기:", font=("맑은 고딕", 12), bg='white').pack(side='left', padx=(0, 5))
        
        # 분기 옵션 생성
        quarters = ["전체", "Q1", "Q2", "Q3", "Q4"]
        
        # 다중 선택을 위한 Listbox로 변경
        quarter_listbox_frame = tk.Frame(quarter_filter_frame, bg='white')
        quarter_listbox_frame.pack(side='left')
        
        self.quarter_listbox = tk.Listbox(quarter_listbox_frame, selectmode='multiple', height=4, width=8)
        self.quarter_listbox.pack(side='left')
        
        # 분기 옵션 추가
        for quarter in quarters:
            self.quarter_listbox.insert(tk.END, quarter)
        
        # 스크롤바
        quarter_scrollbar = tk.Scrollbar(quarter_listbox_frame, orient="vertical", command=self.quarter_listbox.yview)
        self.quarter_listbox.configure(yscrollcommand=quarter_scrollbar.set)
        quarter_scrollbar.pack(side='right', fill='y')
        
        # 비교 필터 프레임 (초기에는 숨김)
        self.compare_filter_frame = tk.Frame(filter_frame, bg='white')
        
        # 비교 대상 1
        compare1_frame = tk.Frame(self.compare_filter_frame, bg='white')
        compare1_frame.pack(side='left', padx=(0, 20))
        
        tk.Label(compare1_frame, text="비교 대상 1:", font=("맑은 고딕", 10), bg='white').pack(side='left', padx=(0, 5))
        
        self.compare1_year = tk.StringVar()
        self.compare1_quarter = tk.StringVar()
        
        compare1_year_combo = ttk.Combobox(compare1_frame, textvariable=self.compare1_year, width=8)
        compare1_year_combo.pack(side='left', padx=(0, 5))
        
        compare1_quarter_combo = ttk.Combobox(compare1_frame, textvariable=self.compare1_quarter, width=5)
        compare1_quarter_combo.pack(side='left')
        
        # 비교 대상 2
        compare2_frame = tk.Frame(self.compare_filter_frame, bg='white')
        compare2_frame.pack(side='left', padx=(0, 20))
        
        tk.Label(compare2_frame, text="비교 대상 2:", font=("맑은 고딕", 10), bg='white').pack(side='left', padx=(0, 5))
        
        self.compare2_year = tk.StringVar()
        self.compare2_quarter = tk.StringVar()
        
        compare2_year_combo = ttk.Combobox(compare2_frame, textvariable=self.compare2_year, width=8)
        compare2_year_combo.pack(side='left', padx=(0, 5))
        
        compare2_quarter_combo = ttk.Combobox(compare2_frame, textvariable=self.compare2_quarter, width=5)
        compare2_quarter_combo.pack(side='left')
        
        # 콤보박스 옵션 설정
        if not self.foundry_df.empty:
            years = sorted(self.foundry_df['FY'].unique().tolist())
        elif not self.foundry_rev_df.empty:
            years = sorted(self.foundry_rev_df['연도'].unique().tolist())
        else:
            years = []
        
        quarters = ["Q1", "Q2", "Q3", "Q4"]
        
        compare1_year_combo['values'] = years
        compare1_quarter_combo['values'] = quarters
        compare2_year_combo['values'] = years
        compare2_quarter_combo['values'] = quarters
        
        # 초기 선택값 설정 (이벤트 바인딩 없이)
        self.selected_years = ["전체"]
        self.selected_quarters = ["전체"]
        
        # 기본값으로 "전체" 선택 (이벤트 발생 없이)
        self.year_listbox.selection_set(0)
        self.quarter_listbox.selection_set(0)
        
        # 이제 이벤트 바인딩 (초기화 완료 후)
        self.year_listbox.bind('<<ListboxSelect>>', self.on_year_changed)
        self.quarter_listbox.bind('<<ListboxSelect>>', self.on_quarter_changed)
        
        print(f"초기 선택값 설정: 연도={self.selected_years}, 분기={self.selected_quarters}")
    
    def on_filter_mode_changed(self):
        """필터 모드 변경 시 호출"""
        if self.filter_mode.get() == "기본":
            self.basic_filter_frame.pack(side='left', padx=(0, 20))
            self.compare_filter_frame.pack_forget()
        else:
            self.basic_filter_frame.pack_forget()
            self.compare_filter_frame.pack(side='left', padx=(0, 20))
        
        # 대시보드 업데이트
        self.update_dashboard()
    
    def on_year_changed(self, event=None):
        """연도 변경 시 호출"""
        if self._updating_year:
            return
        
        self._updating_year = True
        try:
            # 선택된 연도들 가져오기
            selected_indices = self.year_listbox.curselection()
            selected_years = []
            
            for index in selected_indices:
                year = self.year_listbox.get(index)
                selected_years.append(year)
            
            # 선택된 항목이 없으면 "전체"로 설정
            if not selected_years:
                selected_years = ["전체"]
                # 이벤트 바인딩 일시 해제 후 UI 업데이트
                self.year_listbox.unbind('<<ListboxSelect>>')
                self.year_listbox.selection_clear(0, tk.END)
                self.year_listbox.selection_set(0)
                # 약간의 지연 후 이벤트 바인딩 재설정
                self.parent_frame.after(10, lambda: self.year_listbox.bind('<<ListboxSelect>>', self.on_year_changed))
            
            # "전체"가 선택되어 있으면 모든 연도 선택
            if "전체" in selected_years:
                self.selected_years = ["전체"]
            else:
                self.selected_years = selected_years
            
            print(f"연도 선택 업데이트: {self.selected_years}")
            
            # 디바운싱 적용
            self.schedule_update()
        finally:
            self._updating_year = False
    
    def on_quarter_changed(self, event=None):
        """분기 변경 시 호출"""
        if self._updating_quarter:
            return
        
        self._updating_quarter = True
        try:
            # 선택된 분기들 가져오기
            selected_indices = self.quarter_listbox.curselection()
            selected_quarters = []
            
            for index in selected_indices:
                quarter = self.quarter_listbox.get(index)
                selected_quarters.append(quarter)
            
            # 선택된 항목이 없으면 "전체"로 설정
            if not selected_quarters:
                selected_quarters = ["전체"]
                # 이벤트 바인딩 일시 해제 후 UI 업데이트
                self.quarter_listbox.unbind('<<ListboxSelect>>')
                self.quarter_listbox.selection_clear(0, tk.END)
                self.quarter_listbox.selection_set(0)
                # 약간의 지연 후 이벤트 바인딩 재설정
                self.parent_frame.after(10, lambda: self.quarter_listbox.bind('<<ListboxSelect>>', self.on_quarter_changed))
            
            # "전체"가 선택되어 있으면 모든 분기 선택
            if "전체" in selected_quarters:
                self.selected_quarters = ["전체"]
            else:
                self.selected_quarters = selected_quarters
            
            print(f"분기 선택 업데이트: {self.selected_quarters}")
            
            # 디바운싱 적용
            self.schedule_update()
        finally:
            self._updating_quarter = False
    
    def schedule_update(self):
        """디바운싱을 위한 업데이트 스케줄링"""
        if self.update_timer:
            self.parent_frame.after_cancel(self.update_timer)
        
        # 300ms 후에 업데이트 실행
        self.update_timer = self.parent_frame.after(300, self.update_dashboard)
    
    def setup_summary_metrics(self, parent):
        """요약 지표 설정"""
        self.summary_frame = tk.Frame(parent, bg='white')
        self.summary_frame.pack(fill='x', pady=(0, 20))
        
        # 5개의 지표 박스 생성
        metrics = [
            ("목표", "target", "#3498db"),
            ("실적", "actual", "#2ecc71"),
            ("달성율", "achievement", "#f39c12"),
            ("고객사 수", "customers", "#9b59b6"),
            ("YoY", "yoy", "#e74c3c")
        ]
        
        for i, (label, key, color) in enumerate(metrics):
            metric_frame = tk.Frame(self.summary_frame, bg=color, relief="solid", bd=1)
            metric_frame.pack(side='left', fill='both', expand=True, padx=(0 if i == 0 else 10, 0))
            
            # 지표 라벨
            label_widget = tk.Label(metric_frame, text=label, font=("맑은 고딕", 14, "bold"), 
                                   bg=color, fg='white')
            label_widget.pack(pady=(10, 5))
            
            # 지표 값
            value_widget = tk.Label(metric_frame, text="$ 0K", font=("맑은 고딕", 18, "bold"), 
                                   bg=color, fg='white')
            value_widget.pack(pady=(0, 10))
            
            # 위젯 저장
            setattr(self, f"{key}_label", label_widget)
            setattr(self, f"{key}_value", value_widget)
    
    def setup_tabs(self, parent):
        """탭 설정"""
        # 탭 컨트롤
        self.tab_control = ttk.Notebook(parent)
        self.tab_control.pack(fill='both', expand=True)
        
        # 탭 생성
        tabs = [
            ("전체 실적", self.create_overall_performance_tab),
            ("제품별", self.create_product_tab),
            ("라이센스 타입별", self.create_license_type_tab),
            ("고객사별", self.create_customer_tab)
        ]
        
        for tab_name, create_func in tabs:
            tab_frame = tk.Frame(self.tab_control, bg='white')
            self.tab_control.add(tab_frame, text=tab_name)
            create_func(tab_frame)
        
        # 탭 변경 이벤트 바인딩
        self.tab_control.bind('<<NotebookTabChanged>>', self.on_tab_changed)
    
    def create_overall_performance_tab(self, parent):
        """전체 실적 탭 생성"""
        # 인터랙티브 차트 섹션
        chart_label = tk.Label(parent, text="인터랙티브 차트", font=("맑은 고딕", 14, "bold"), 
                              bg='white', fg='#2c3e50')
        chart_label.pack(anchor='w', pady=(10, 5))
        
        # 차트와 테이블을 담을 프레임
        content_frame = tk.Frame(parent, bg='white')
        content_frame.pack(fill='both', expand=True)
        
        # 테이블 프레임 (왼쪽)
        self.table_frame = tk.Frame(content_frame, bg='white')
        self.table_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        # 차트 프레임 (오른쪽)
        self.chart_frame = tk.Frame(content_frame, bg='white')
        self.chart_frame.pack(side='right', fill='both', expand=True)
    
    def create_product_tab(self, parent):
        """제품별 탭 생성"""
        # 제품별 차트와 테이블을 담을 프레임
        content_frame = tk.Frame(parent, bg='white')
        content_frame.pack(fill='both', expand=True)
        
        # 테이블 프레임 (왼쪽)
        self.product_table_frame = tk.Frame(content_frame, bg='white')
        self.product_table_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        # 차트 프레임 (오른쪽)
        self.product_chart_frame = tk.Frame(content_frame, bg='white')
        self.product_chart_frame.pack(side='right', fill='both', expand=True)
    
    def create_license_type_tab(self, parent):
        """라이센스 타입별 탭 생성"""
        # 라이센스 타입별 차트와 테이블을 담을 프레임
        content_frame = tk.Frame(parent, bg='white')
        content_frame.pack(fill='both', expand=True)
        
        # 테이블 프레임 (왼쪽)
        self.license_table_frame = tk.Frame(content_frame, bg='white')
        self.license_table_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        # 차트 프레임 (오른쪽)
        self.license_chart_frame = tk.Frame(content_frame, bg='white')
        self.license_chart_frame.pack(side='right', fill='both', expand=True)
    
    def create_customer_tab(self, parent):
        """고객사별 탭 생성"""
        # 고객사별 테이블만 표시
        self.customer_table_frame = tk.Frame(parent, bg='white')
        self.customer_table_frame.pack(fill='both', expand=True)
    
    def on_tab_changed(self, event=None):
        """탭 변경 시 호출"""
        if self.tab_control:
            current_tab = self.tab_control.tab(self.tab_control.select(), "text")
            self.current_tab = current_tab
            self.update_dashboard()
    
    def update_dashboard(self):
        """대시보드 업데이트"""
        if self._updating_dashboard:
            return
        
        self._updating_dashboard = True
        try:
            # 요약 지표 업데이트
            self.update_summary_metrics()
            
            # 현재 탭에 따른 업데이트
            if self.current_tab == "전체 실적":
                self.update_overall_performance()
            elif self.current_tab == "제품별":
                self.update_product_tab()
            elif self.current_tab == "라이센스 타입별":
                self.update_license_type_tab()
            elif self.current_tab == "고객사별":
                self.update_customer_tab()
                
        finally:
            self._updating_dashboard = False
    
    def update_summary_metrics(self):
        """요약 지표 업데이트"""
        try:
            # 필터링된 데이터 생성
            filtered_df = self.get_filtered_data()
            
            # 목표 계산
            target_value = self.calculate_target()
            
            # 실적 계산
            actual_value = filtered_df['Revenue'].sum() if not filtered_df.empty else 0
            
            # 달성율 계산
            achievement_rate = (actual_value / target_value * 100) if target_value > 0 else 0
            
            # 고객사 수 계산 - Account 컬럼 기준으로 중복 제거
            customer_count = len(filtered_df['Account'].unique()) if not filtered_df.empty else 0
            
            # YoY 계산 (간단한 예시)
            yoy_value = 0  # 실제로는 전년 대비 계산 필요
            
            # 지표 업데이트 - 천단위 K로 표시
            self.target_value.config(text=f"$ {target_value/1000:,.0f}K")
            self.actual_value.config(text=f"$ {actual_value/1000:,.0f}K")
            self.achievement_value.config(text=f"{achievement_rate:.1f}%")
            self.customers_value.config(text=f"{customer_count:,}개")
            self.yoy_value.config(text=f"{yoy_value:+.1f}%")
            
        except Exception as e:
            print(f"요약 지표 업데이트 실패: {e}")
    
    def get_filtered_data(self):
        """필터링된 데이터 반환"""
        try:
            if self.foundry_df is None or self.foundry_df.empty:
                return pd.DataFrame()
        
            # 원본 데이터 복사
            filtered_df = self.foundry_df.copy()
            print(f"원본 데이터 크기: {len(filtered_df)}행")
            
            # 필터 모드에 따라 다른 처리
            if self.filter_mode.get() == "비교":
                # 비교 모드: 두 개의 특정 연도/분기 조합 비교
                compare_data = self.get_compare_data()
                if not compare_data.empty:
                    return compare_data
                else:
                    # 비교 데이터가 없으면 기본 필터 사용
                    pass
            
            # 기본 모드 또는 비교 데이터가 없는 경우
            print(f"선택된 연도: {self.selected_years}")
            print(f"선택된 분기: {self.selected_quarters}")
            
            # 연도 필터링
            if "전체" not in self.selected_years:
                print(f"연도 필터 적용 전: {len(filtered_df)}행")
                print(f"사용 가능한 연도: {filtered_df['FY'].unique() if 'FY' in filtered_df.columns else 'FY 컬럼 없음'}")
                filtered_df = filtered_df[filtered_df['FY'].isin(self.selected_years)]
                print(f"연도 필터 적용 후: {len(filtered_df)}행")
            
            # 분기 필터링
            if "전체" not in self.selected_quarters:
                print(f"분기 필터 적용 전: {len(filtered_df)}행")
                print(f"사용 가능한 분기: {filtered_df['Quarter'].unique() if 'Quarter' in filtered_df.columns else 'Quarter 컬럼 없음'}")
                filtered_df = filtered_df[filtered_df['Quarter'].isin(self.selected_quarters)]
                print(f"분기 필터 적용 후: {len(filtered_df)}행")
            
            print(f"최종 필터링된 데이터 크기: {len(filtered_df)}행")
            if not filtered_df.empty:
                print(f"필터링된 데이터 샘플:")
                print(filtered_df.head(3))
            
            return filtered_df
            
        except Exception as e:
            print(f"데이터 필터링 실패: {e}")
            return pd.DataFrame()
    
    def get_compare_data(self):
        """비교 모드용 데이터 반환"""
        try:
            if self.foundry_df is None or self.foundry_df.empty:
                return pd.DataFrame()
            
            # 비교 대상 1과 2의 데이터 가져오기
            compare1_year = self.compare1_year.get()
            compare1_quarter = self.compare1_quarter.get()
            compare2_year = self.compare2_year.get()
            compare2_quarter = self.compare2_quarter.get()
            
            if not all([compare1_year, compare1_quarter, compare2_year, compare2_quarter]):
                print("비교 대상이 모두 선택되지 않았습니다.")
                return pd.DataFrame()
            
            # 두 비교 대상의 데이터 필터링
            compare1_data = self.foundry_df[
                (self.foundry_df['FY'] == compare1_year) & 
                (self.foundry_df['Quarter'] == compare1_quarter)
            ].copy()
            
            compare2_data = self.foundry_df[
                (self.foundry_df['FY'] == compare2_year) & 
                (self.foundry_df['Quarter'] == compare2_quarter)
            ].copy()
            
            # 비교 구분을 위한 컬럼 추가
            compare1_data['Compare_Group'] = f"{compare1_year} {compare1_quarter}"
            compare2_data['Compare_Group'] = f"{compare2_year} {compare2_quarter}"
            
            # 데이터 합치기
            combined_data = pd.concat([compare1_data, compare2_data], ignore_index=True)
            
            print(f"비교 데이터: {compare1_year} {compare1_quarter} ({len(compare1_data)}행) vs {compare2_year} {compare2_quarter} ({len(compare2_data)}행)")
            print(f"총 비교 데이터: {len(combined_data)}행")
            
            return combined_data
            
        except Exception as e:
            print(f"비교 데이터 생성 실패: {e}")
            return pd.DataFrame()
    
    def calculate_target(self):
        """목표 계산"""
        try:
            if self.foundry_rev_df.empty:
                return 0
            
            # 현재 선택된 연도와 분기에 따른 목표 계산
            target_value = 0
            
            for _, row in self.foundry_rev_df.iterrows():
                year = row['연도']
                quarter_col = None
                
                # 분기 컬럼 결정
                if "전체" in self.selected_quarters:
                    quarter_col = 'TTL'
                elif "Q1" in self.selected_quarters:
                    quarter_col = 'Q1'
                elif "Q2" in self.selected_quarters:
                    quarter_col = 'Q2'
                elif "Q3" in self.selected_quarters:
                    quarter_col = 'Q3'
                elif "Q4" in self.selected_quarters:
                    quarter_col = 'Q4'
                
                # 연도 필터
                if "전체" in self.selected_years or year in self.selected_years:
                    if quarter_col and quarter_col in row:
                        target_value += row[quarter_col] if pd.notna(row[quarter_col]) else 0
            
            return target_value
            
        except Exception as e:
            print(f"목표 계산 실패: {e}")
            return 0
    
    def update_overall_performance(self):
        """전체 실적 탭 업데이트"""
        try:
            # 기존 위젯 정리
            for widget in self.table_frame.winfo_children():
                widget.destroy()
            
            for widget in self.chart_frame.winfo_children():
                widget.destroy()
            
            # 필터링된 데이터
            filtered_df = self.get_filtered_data()
            
            if filtered_df.empty:
                # 데이터가 없을 때 메시지 표시
                no_data_label = tk.Label(self.table_frame, text="선택된 조건에 해당하는 데이터가 없습니다.",
                                        font=("맑은 고딕", 12), bg='white', fg='#7f8c8d')
                no_data_label.pack(expand=True)
                return
            
            # 테이블 생성 (필터링된 데이터 전달)
            self.create_performance_table(self.table_frame, filtered_df)
            
            # 차트 생성 (필터링된 데이터 전달)
            self.create_performance_chart(self.chart_frame, filtered_df)
            
        except Exception as e:
            print(f"전체 실적 업데이트 실패: {e}")
    
    def create_performance_table(self, parent, filtered_df=None):
        """전체 실적 테이블 생성 - 필터링된 데이터 반영"""
        print("=== 전체 실적 테이블 생성 시작 ===")
        
        # 테이블 생성
        columns = ['연도', '구분', 'Q1', 'Q2', 'Q3', 'Q4', 'TTL']
        
        tree = ttk.Treeview(parent, columns=columns, show='headings', height=25)
        print(f"테이블 생성됨: {columns}")
        
        # 컬럼 설정
        for col in columns:
            tree.heading(col, text=col)
            if col == '연도':
                tree.column(col, width=80, anchor='w')
            elif col == '구분':
                tree.column(col, width=80, anchor='w')
            else:
                tree.column(col, width=120, anchor='e')
        
        # 스크롤바
        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        print("테이블 UI 구성 완료")
        
        # 목표 데이터 (Foundry_Rev에서) - 필터링 적용
        target_data = {}
        if hasattr(self, 'foundry_rev_df') and self.foundry_rev_df is not None:
            rev_data = self.foundry_rev_df
            print(f"Foundry_Rev 데이터 로드됨: {len(rev_data)}행")
            
            # 연도 필터 적용
            if "전체" not in self.selected_years:
                rev_data = rev_data[rev_data['연도'].isin(self.selected_years)]
                print(f"목표 데이터 연도 필터 적용 후: {len(rev_data)}행")
            
            for _, row in rev_data.iterrows():
                year = row['연도']
                category = row['구분']
                if category == '목표':
                    target_data[year] = {
                        'Q1': row['Q1'] if pd.notna(row['Q1']) else 0,
                        'Q2': row['Q2'] if pd.notna(row['Q2']) else 0,
                        'Q3': row['Q3'] if pd.notna(row['Q3']) else 0,
                        'Q4': row['Q4'] if pd.notna(row['Q4']) else 0,
                        'TTL': row['TTL'] if pd.notna(row['TTL']) else 0
                    }
        
        # 실적 데이터 (필터링된 데이터 사용)
        actual_data = {}
        if filtered_df is not None and not filtered_df.empty:
            print(f"필터링된 데이터 사용: {len(filtered_df)}행")
            
            # 연도별, 분기별 실적 집계 (필터링된 데이터에서)
            for year in filtered_df['FY'].unique():
                year_data = filtered_df[filtered_df['FY'] == year]
                if not year_data.empty:
                    actual_data[year] = {
                        'Q1': year_data[year_data['Quarter'] == 'Q1']['Revenue'].sum() if 'Revenue' in year_data.columns else 0,
                        'Q2': year_data[year_data['Quarter'] == 'Q2']['Revenue'].sum() if 'Revenue' in year_data.columns else 0,
                        'Q3': year_data[year_data['Quarter'] == 'Q3']['Revenue'].sum() if 'Revenue' in year_data.columns else 0,
                        'Q4': year_data[year_data['Quarter'] == 'Q4']['Revenue'].sum() if 'Revenue' in year_data.columns else 0,
                        'TTL': year_data['Revenue'].sum() if 'Revenue' in year_data.columns else 0
                    }
                else:
                    actual_data[year] = {'Q1': 0, 'Q2': 0, 'Q3': 0, 'Q4': 0, 'TTL': 0}
        else:
            # 기존 방식 (하위 호환성)
            if hasattr(self, 'foundry_df') and self.foundry_df is not None:
                foundry_data = self.foundry_df
                print(f"기존 방식으로 Foundry 데이터 사용: {len(foundry_data)}행")
                
                # 연도별, 분기별 실적 집계
                for year in ['FY20', 'FY21', 'FY22', 'FY23', 'FY24', 'FY25']:
                    year_data = foundry_data[foundry_data['FY'] == year]
                    if not year_data.empty:
                        actual_data[year] = {
                            'Q1': year_data[year_data['Quarter'] == 'Q1']['Unit Price'].sum() if 'Unit Price' in year_data.columns else 0,
                            'Q2': year_data[year_data['Quarter'] == 'Q2']['Unit Price'].sum() if 'Unit Price' in year_data.columns else 0,
                            'Q3': year_data[year_data['Quarter'] == 'Q3']['Unit Price'].sum() if 'Unit Price' in year_data.columns else 0,
                            'Q4': year_data[year_data['Quarter'] == 'Q4']['Unit Price'].sum() if 'Unit Price' in year_data.columns else 0,
                            'TTL': year_data['Unit Price'].sum() if 'Unit Price' in year_data.columns else 0
                        }
                    else:
                        actual_data[year] = {'Q1': 0, 'Q2': 0, 'Q3': 0, 'Q4': 0, 'TTL': 0}
        
        # 모든 연도 처리
        all_years = sorted(set(list(target_data.keys()) + list(actual_data.keys())))
        print(f"처리할 연도: {all_years}")
        
        for year in all_years:
            print(f"\n=== {year} 처리 중 ===")
            
            # 목표 데이터
            target = target_data.get(year, {'Q1': 0, 'Q2': 0, 'Q3': 0, 'Q4': 0, 'TTL': 0})
            print(f"{year} 목표: {target}")
            
            # 실적 데이터
            actual = actual_data.get(year, {'Q1': 0, 'Q2': 0, 'Q3': 0, 'Q4': 0, 'TTL': 0})
            print(f"{year} 실적: {actual}")
            
            # 목표 행 추가
            target_item = tree.insert('', 'end', values=[
                year, '목표',
                f"{target['Q1']:,.2f}" if target['Q1'] > 0 else '',
                f"{target['Q2']:,.2f}" if target['Q2'] > 0 else '',
                f"{target['Q3']:,.2f}" if target['Q3'] > 0 else '',
                f"{target['Q4']:,.2f}" if target['Q4'] > 0 else '',
                f"{target['TTL']:,.2f}" if target['TTL'] > 0 else ''
            ], tags=('target',))
            print(f"{year} 목표 행 추가됨")
            
            # 실적 행 추가
            actual_item = tree.insert('', 'end', text='', values=[
                '', '실적',
                f"{actual['Q1']:,.2f}" if actual['Q1'] > 0 else '',
                f"{actual['Q2']:,.2f}" if actual['Q2'] > 0 else '',
                f"{actual['Q3']:,.2f}" if actual['Q3'] > 0 else '',
                f"{actual['Q4']:,.2f}" if actual['Q4'] > 0 else '',
                f"{actual['TTL']:,.2f}" if actual['TTL'] > 0 else ''
            ], tags=('actual',))
            print(f"{year} 실적 행 추가됨")
            
            # 달성율 계산 및 추가
            achievement = {}
            for quarter in ['Q1', 'Q2', 'Q3', 'Q4', 'TTL']:
                if target[quarter] > 0:
                    achievement[quarter] = (actual[quarter] / target[quarter]) * 100
                else:
                    achievement[quarter] = 0
            
            tree.insert('', 'end', text='', values=[
                '', '달성율',
                f"{achievement['Q1']:.1f}%" if achievement['Q1'] > 0 else '',
                f"{achievement['Q2']:.1f}%" if achievement['Q2'] > 0 else '',
                f"{achievement['Q3']:.1f}%" if achievement['Q3'] > 0 else '',
                f"{achievement['Q4']:.1f}%" if achievement['Q4'] > 0 else '',
                f"{achievement['TTL']:.1f}%" if achievement['TTL'] > 0 else ''
            ], tags=('achievement',))
            print(f"{year} 달성율 행 추가됨")
            
            # YoY 계산 및 추가
            if year != all_years[0]:  # 첫 번째 연도가 아닌 경우
                prev_year = all_years[all_years.index(year) - 1]
                prev_actual = actual_data.get(prev_year, {'Q1': 0, 'Q2': 0, 'Q3': 0, 'Q4': 0, 'TTL': 0})
                
                yoy = {}
                for quarter in ['Q1', 'Q2', 'Q3', 'Q4', 'TTL']:
                    if prev_actual[quarter] > 0:
                        yoy[quarter] = ((actual[quarter] - prev_actual[quarter]) / prev_actual[quarter]) * 100
                    else:
                        yoy[quarter] = 0
                
                tree.insert('', 'end', text='', values=[
                    '', 'YoY',
                    f"{yoy['Q1']:.1f}%" if yoy['Q1'] != 0 else '',
                    f"{yoy['Q2']:.1f}%" if yoy['Q2'] != 0 else '',
                    f"{yoy['Q3']:.1f}%" if yoy['Q3'] != 0 else '',
                    f"{yoy['Q4']:.1f}%" if yoy['Q4'] != 0 else '',
                    f"{yoy['TTL']:.1f}%" if yoy['TTL'] != 0 else ''
                ], tags=('yoy',))
                print(f"{year} YoY 행 추가됨")
            else:
                # 첫 번째 연도의 경우 YoY는 비어있음
                tree.insert('', 'end', text='', values=['', 'YoY', '', '', '', '', ''], tags=('yoy',))
                print(f"{year} YoY 빈 행 추가됨 (첫 번째 연도)")
        
        # 태그별 스타일 설정
        tree.tag_configure('target', background='#f0f0f0')
        tree.tag_configure('actual', background='#e6f3ff')
        tree.tag_configure('achievement', background='#fff2e6')
        tree.tag_configure('yoy', background='#f0fff0')
        
        # 더블클릭 이벤트 추가
        tree.bind('<Double-1>', lambda e: self.on_performance_double_click(e, tree, filtered_df))
        
        print("=== 전체 실적 테이블 생성 완료 ===")
        return tree
    
    def on_performance_double_click(self, event, tree, filtered_df):
        """전체 실적 테이블 더블클릭 이벤트"""
        item = tree.identify_row(event.y)
        if not item:
            return
        
        values = tree.item(item, 'values')
        if not values or len(values) < 2:
            return
        
        year = values[0]
        category = values[1]
        
        # 연도가 비어있으면 상위 행에서 찾기
        if not year:
            parent = tree.parent(item)
            if parent:
                parent_values = tree.item(parent, 'values')
                if parent_values and len(parent_values) > 0:
                    year = parent_values[0]
        
        if not year:
            return
        
        # 해당 연도/분기 데이터 필터링
        if filtered_df is not None and not filtered_df.empty:
            year_data = filtered_df[filtered_df['FY'] == year].copy()
            
            if category == '실적':
                # 실적 데이터만 표시
                self.show_performance_detail(year, year_data, "실적")
            elif category in ['Q1', 'Q2', 'Q3', 'Q4']:
                # 특정 분기 데이터만 표시
                quarter_data = year_data[year_data['Quarter'] == category].copy()
                self.show_performance_detail(year, quarter_data, f"{category} 실적")
            elif category == 'TTL':
                # 전체 연도 데이터 표시
                self.show_performance_detail(year, year_data, "전체 실적")
    
    def show_performance_detail(self, year, data, title_suffix):
        """실적 상세 데이터 창 표시"""
        if data.empty:
            messagebox.showinfo("알림", f"{year} {title_suffix} 데이터가 없습니다.")
            return
        
        # 상세 데이터 창 생성
        detail_window = tk.Toplevel(self.parent)
        detail_window.title(f"{year} {title_suffix} - 상세 데이터")
        detail_window.geometry("1200x800")
        detail_window.configure(bg="#F7F9FB")
        
        # 헤더
        header_frame = tk.Frame(detail_window, bg="#F7F9FB", height=60)
        header_frame.pack(fill='x', pady=(0, 20))
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(header_frame, text=f"📊 {year} {title_suffix} - 상세 데이터", 
                              font=("맑은 고딕", 16, "bold"), 
                              fg="#1F2937", bg="#F7F9FB")
        title_label.pack(side='left', padx=20, pady=15)
        
        # 통계 정보
        stats_frame = tk.Frame(detail_window, bg="white", relief="solid", bd=1)
        stats_frame.pack(fill='x', padx=20, pady=(0, 20))
        
        # 통계 카드들
        total_revenue = data['Revenue'].sum() if 'Revenue' in data.columns else 0
        total_count = len(data)
        avg_revenue = total_revenue / total_count if total_count > 0 else 0
        
        stats_cards = [
            ("총 매출", f"{total_revenue:,.2f}"),
            ("거래 건수", f"{total_count:,}건"),
            ("평균 매출", f"{avg_revenue:,.2f}")
        ]
        
        for i, (label, value) in enumerate(stats_cards):
            card = tk.Frame(stats_frame, bg="#F8FAFC", relief="solid", bd=1)
            card.pack(side='left', fill='both', expand=True, padx=5, pady=10)
            
            tk.Label(card, text=label, font=("맑은 고딕", 10), 
                    fg="#6B7280", bg="#F8FAFC").pack(pady=(10, 5))
            tk.Label(card, text=value, font=("맑은 고딕", 14, "bold"), 
                    fg="#1F2937", bg="#F8FAFC").pack(pady=(0, 10))
        
        # 테이블 프레임
        table_frame = tk.Frame(detail_window, bg="white", relief="solid", bd=1)
        table_frame.pack(fill='both', expand=True, padx=20, pady=(0, 20))
        
        # 사용 가능한 컬럼 선택
        available_columns = [col for col in data.columns if col not in ['FY', 'Quarter']]
        if len(available_columns) > 10:  # 컬럼이 너무 많으면 주요 컬럼만 선택
            priority_columns = ['Date', 'Customer', 'Product', 'Revenue', 'Unit Price', 'Quantity']
            available_columns = [col for col in priority_columns if col in data.columns]
            available_columns.extend([col for col in data.columns if col not in priority_columns and col not in available_columns][:5])
        
        # 트리뷰
        tree = ttk.Treeview(table_frame, columns=available_columns, show='headings', height=20)
        
        # 스크롤바
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # 컬럼 설정
        for col in available_columns:
            tree.heading(col, text=col)
            if col in ['Revenue', 'Unit Price', 'Quantity']:
                tree.column(col, width=120, anchor='e')
            else:
                tree.column(col, width=150, anchor='w')
        
        # 데이터 삽입
        for _, row in data.iterrows():
            values = []
            for col in available_columns:
                val = row.get(col, '')
                if col in ['Revenue', 'Unit Price']:
                    try:
                        val = f"{float(val):,.2f}"
                    except:
                        val = str(val)
                elif col == 'Quantity':
                    try:
                        val = f"{int(val):,}"
                    except:
                        val = str(val)
                values.append(val)
            tree.insert('', 'end', values=values)
        
        # 배치
        tree.pack(side='left', fill='both', expand=True, padx=10, pady=10)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
    
    def create_performance_chart(self, parent, data):
        """실적 차트 생성 - 필터링된 데이터 반영"""
        try:
            # 차트 제목
            title_label = tk.Label(parent, text="연도별 목표 vs 실적", font=("맑은 고딕", 12, "bold"), 
                                  bg='white', fg='#2c3e50')
            title_label.pack(anchor='w', pady=(0, 10))
            
            # matplotlib 차트 생성
            fig, ax = plt.subplots(figsize=(10, 6))
            
            # 필터 모드에 따라 다른 차트 생성
            if self.filter_mode.get() == "비교":
                self.create_compare_chart(ax, data)
            else:
                self.create_normal_chart(ax, data)
            
            plt.tight_layout()
            
            # Tkinter에 차트 삽입
            canvas = FigureCanvasTkAgg(fig, parent)
            canvas.draw()
            canvas.get_tk_widget().pack(fill='both', expand=True)
            
        except Exception as e:
            print(f"차트 생성 실패: {e}")
            import traceback
            traceback.print_exc()
            # 오류 시 메시지 표시
            error_label = tk.Label(parent, text=f"차트 생성 중 오류가 발생했습니다.\n{e}",
                                  font=("맑은 고딕", 10), bg='white', fg='#e74c3c')
            error_label.pack(expand=True)
    
    def create_normal_chart(self, ax, data):
        """일반 모드 차트 생성"""
        # 연도별 데이터 준비 (필터링된 데이터 사용)
        if data is not None and not data.empty:
            # 필터링된 데이터에서 연도 추출
            years = sorted(data['FY'].unique())
            print(f"차트용 연도: {years}")
        else:
            # 기존 방식 (하위 호환성)
            years = sorted(self.foundry_rev_df['연도'].unique()) if not self.foundry_rev_df.empty else []
        
        if years:
            # 분기별 색상 설정
            quarter_colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728']  # Q1, Q2, Q3, Q4
            
            # 목표와 실적 데이터 준비 (필터링된 데이터 사용)
            target_data = []
            actual_data = []
            year_labels = []
            
            for year in years:
                # 목표 데이터 (Foundry_Rev에서)
                target_row = self.foundry_rev_df[(self.foundry_rev_df['연도'] == year) & 
                                               (self.foundry_rev_df['구분'] == '목표')]
                
                # 실적 데이터 (필터링된 데이터에서 계산)
                if data is not None and not data.empty:
                    year_data = data[data['FY'] == year]
                    if not year_data.empty:
                        actual_q1 = year_data[year_data['Quarter'] == 'Q1']['Revenue'].sum() if 'Revenue' in year_data.columns else 0
                        actual_q2 = year_data[year_data['Quarter'] == 'Q2']['Revenue'].sum() if 'Revenue' in year_data.columns else 0
                        actual_q3 = year_data[year_data['Quarter'] == 'Q3']['Revenue'].sum() if 'Revenue' in year_data.columns else 0
                        actual_q4 = year_data[year_data['Quarter'] == 'Q4']['Revenue'].sum() if 'Revenue' in year_data.columns else 0
                    else:
                        actual_q1 = actual_q2 = actual_q3 = actual_q4 = 0
                else:
                    # 기존 방식 (하위 호환성)
                    actual_row = self.foundry_rev_df[(self.foundry_rev_df['연도'] == year) & 
                                                   (self.foundry_rev_df['구분'] == '실적')]
                    if not actual_row.empty:
                        actual_q1 = actual_row['Q1'].iloc[0] if pd.notna(actual_row['Q1'].iloc[0]) else 0
                        actual_q2 = actual_row['Q2'].iloc[0] if pd.notna(actual_row['Q2'].iloc[0]) else 0
                        actual_q3 = actual_row['Q3'].iloc[0] if pd.notna(actual_row['Q3'].iloc[0]) else 0
                        actual_q4 = actual_row['Q4'].iloc[0] if pd.notna(actual_row['Q4'].iloc[0]) else 0
                    else:
                        actual_q1 = actual_q2 = actual_q3 = actual_q4 = 0
                
                if not target_row.empty:
                    target_q1 = target_row['Q1'].iloc[0] if pd.notna(target_row['Q1'].iloc[0]) else 0
                    target_q2 = target_row['Q2'].iloc[0] if pd.notna(target_row['Q2'].iloc[0]) else 0
                    target_q3 = target_row['Q3'].iloc[0] if pd.notna(target_row['Q3'].iloc[0]) else 0
                    target_q4 = target_row['Q4'].iloc[0] if pd.notna(target_row['Q4'].iloc[0]) else 0
                    target_data.append([target_q1, target_q2, target_q3, target_q4])
                else:
                    target_data.append([0, 0, 0, 0])
                
                actual_data.append([actual_q1, actual_q2, actual_q3, actual_q4])
                year_labels.append(year)
            
            # x축 위치 설정
            x = np.arange(len(years))
            width = 0.35
            
            # 목표 바 차트 (스택)
            target_bars = []
            for i, year_data in enumerate(target_data):
                bottom = 0
                for j, quarter_value in enumerate(year_data):
                    if quarter_value > 0:
                        bar = ax.bar(x[i] - width/2, quarter_value, width, 
                                   bottom=bottom, color=quarter_colors[j], 
                                   alpha=0.7, label=f'Q{j+1}' if i == 0 else "")
                        bottom += quarter_value
                        target_bars.append(bar)
            
            # 실적 바 차트 (스택)
            actual_bars = []
            for i, year_data in enumerate(actual_data):
                bottom = 0
                for j, quarter_value in enumerate(year_data):
                    if quarter_value > 0:
                        bar = ax.bar(x[i] + width/2, quarter_value, width, 
                                   bottom=bottom, color=quarter_colors[j], 
                                   alpha=0.9, label=f'Q{j+1}' if i == 0 else "")
                        bottom += quarter_value
                        actual_bars.append(bar)
            
            # 차트 설정
            ax.set_xlabel('Fiscal Year', fontsize=12)
            ax.set_ylabel('Revenue ($)', fontsize=12)
            ax.set_title('연도별 목표 vs 실적 (분기별)', fontsize=14, fontweight='bold')
            ax.set_xticks(x)
            ax.set_xticklabels(year_labels)
            ax.legend(['Q1', 'Q2', 'Q3', 'Q4'], loc='upper left')
            ax.grid(True, alpha=0.3)
            
            # 범례 추가
            from matplotlib.patches import Patch
            legend_elements = [
                Patch(facecolor='gray', alpha=0.7, label='목표'),
                Patch(facecolor='gray', alpha=0.9, label='실적')
            ]
            ax.legend(handles=legend_elements, loc='upper right')
            
            # 값 표시 (총합)
            for i, (target_year, actual_year) in enumerate(zip(target_data, actual_data)):
                target_total = sum(target_year)
                actual_total = sum(actual_year)
                
                if target_total > 0:
                    ax.text(x[i] - width/2, target_total + target_total*0.01, 
                           f'{target_total:,.0f}', ha='center', va='bottom', fontsize=8)
                
                if actual_total > 0:
                    ax.text(x[i] + width/2, actual_total + actual_total*0.01, 
                           f'{actual_total:,.0f}', ha='center', va='bottom', fontsize=8)
            
        else:
            ax.text(0.5, 0.5, '데이터가 없습니다', ha='center', va='center', 
                   transform=ax.transAxes, fontsize=14)
    
    def create_compare_chart(self, ax, data):
        """비교 모드 차트 생성"""
        if data is None or data.empty or 'Compare_Group' not in data.columns:
            ax.text(0.5, 0.5, '비교 데이터가 없습니다', ha='center', va='center', 
                   transform=ax.transAxes, fontsize=14)
            return
        
        # 비교 그룹별 데이터 집계
        compare_groups = data['Compare_Group'].unique()
        if len(compare_groups) != 2:
            ax.text(0.5, 0.5, '비교 대상이 2개가 아닙니다', ha='center', va='center', 
                   transform=ax.transAxes, fontsize=14)
            return
        
        group1, group2 = compare_groups
        
        # 선택된 분기 확인
        selected_quarter = None
        if hasattr(self, 'compare1_quarter') and self.compare1_quarter.get():
            selected_quarter = self.compare1_quarter.get()
        
        # 각 그룹의 목표와 실적 데이터 계산
        group1_data = data[data['Compare_Group'] == group1]
        group2_data = data[data['Compare_Group'] == group2]
        
        # 목표 데이터 (Foundry_Rev에서)
        group1_year = group1.split()[0]  # "FY24 Q2" -> "FY24"
        group2_year = group2.split()[0]  # "FY25 Q2" -> "FY25"
        
        # 목표 데이터 가져오기
        target1_row = self.foundry_rev_df[(self.foundry_rev_df['연도'] == group1_year) & 
                                         (self.foundry_rev_df['구분'] == '목표')]
        target2_row = self.foundry_rev_df[(self.foundry_rev_df['연도'] == group2_year) & 
                                         (self.foundry_rev_df['구분'] == '목표')]
        
        # 선택된 분기의 목표값 (선택한 연도/분기의 목표만 표시)
        target1_value = 0
        target2_value = 0
        
        if not target1_row.empty and selected_quarter:
            target1_value = target1_row[selected_quarter].iloc[0] if pd.notna(target1_row[selected_quarter].iloc[0]) else 0
        
        if not target2_row.empty and selected_quarter:
            target2_value = target2_row[selected_quarter].iloc[0] if pd.notna(target2_row[selected_quarter].iloc[0]) else 0
        
        # 실적 데이터
        actual1_value = group1_data['Revenue'].sum()
        actual2_value = group2_data['Revenue'].sum()
        
        # 차트 데이터 준비
        x = np.arange(2)
        width = 0.35
        
        # 목표 바 (왼쪽)
        ax.bar(x - width/2, [target1_value, target2_value], width, 
               label='목표', color='lightblue', alpha=0.7)
        
        # 실적 바 (오른쪽)
        ax.bar(x + width/2, [actual1_value, actual2_value], width, 
               label='실적', color='lightgreen', alpha=0.9)
        
        # 차트 설정
        ax.set_xlabel('비교 대상', fontsize=12)
        ax.set_ylabel('Revenue ($)', fontsize=12)
        ax.set_title(f'{selected_quarter} 선택 연도/분기 목표 vs 실적 비교', fontsize=14, fontweight='bold')
        ax.set_xticks(x)
        ax.set_xticklabels([group1, group2])
        ax.legend()
        ax.grid(True, alpha=0.3)
        
        # 값 표시 (목표값과 실적값)
        if target1_value > 0:
            ax.text(x[0] - width/2, target1_value + target1_value*0.01, 
                   f'목표: {target1_value:,.0f}', ha='center', va='bottom', fontsize=8, fontweight='bold')
        
        if target2_value > 0:
            ax.text(x[1] - width/2, target2_value + target2_value*0.01, 
                   f'목표: {target2_value:,.0f}', ha='center', va='bottom', fontsize=8, fontweight='bold')
        
        if actual1_value > 0:
            ax.text(x[0] + width/2, actual1_value + actual1_value*0.01, 
                   f'실적: {actual1_value:,.0f}', ha='center', va='bottom', fontsize=8, fontweight='bold')
        
        if actual2_value > 0:
            ax.text(x[1] + width/2, actual2_value + actual2_value*0.01, 
                   f'실적: {actual2_value:,.0f}', ha='center', va='bottom', fontsize=8, fontweight='bold')
    
    def show_detail_window(self, data):
        """상세 정보 창 표시"""
        try:
            DetailWindow(self.parent_frame, "Foundry 상세 정보", data, self.foundry_df)
        except Exception as e:
            print(f"상세 정보 창 표시 실패: {e}")
    
    def update_product_tab(self):
        """제품별 탭 업데이트 - 이미지와 같은 형태"""
        try:
            # 기존 위젯 정리
            for widget in self.product_table_frame.winfo_children():
                widget.destroy()
            
            for widget in self.product_chart_frame.winfo_children():
                widget.destroy()
            
            # 필터링된 데이터
            filtered_df = self.get_filtered_data()
            
            if filtered_df.empty:
                # 데이터가 없을 때 메시지 표시
                no_data_label = tk.Label(self.product_table_frame, text="선택된 조건에 해당하는 데이터가 없습니다.",
                                        font=("맑은 고딕", 12), bg='white', fg='#7f8c8d')
                no_data_label.pack(expand=True)
                return
            
            # 제품별 테이블 생성
            self.create_product_table(self.product_table_frame, filtered_df)
            
            # 제품별 차트 생성
            self.create_product_chart(self.product_chart_frame, filtered_df)
            
        except Exception as e:
            print(f"제품별 탭 업데이트 실패: {e}")
    
    def create_product_table(self, parent, data):
        """제품별 테이블 생성 - 필터링된 데이터 반영"""
        # 테이블 생성
        columns = ['제품명', 'FY20', 'FY21', 'FY22', 'FY23', 'FY24', 'FY25', '총합계']
        
        tree = ttk.Treeview(parent, columns=columns, show='headings', height=20)
        
        # 컬럼 설정
        for col in columns:
            tree.heading(col, text=col)
            if col == '제품명':
                tree.column(col, width=200, anchor='w')
            else:
                tree.column(col, width=100, anchor='e')
        
        # 스크롤바
        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        
        # 제품별 데이터 집계 (필터링된 데이터 사용)
        if 'Product' in data.columns and 'Revenue' in data.columns:
            # 제품별로 데이터 집계
            product_data = data.groupby('Product').agg({
                'Revenue': 'sum'
            }).reset_index()
            
            # 연도별 데이터 집계
            for fy in ['FY20', 'FY21', 'FY22', 'FY23', 'FY24', 'FY25']:
                fy_data = data[data['FY'] == fy].groupby('Product').agg({
                    'Revenue': 'sum'
                }).reset_index()
                
                # 제품별 데이터에 연도별 데이터 추가
                for _, row in fy_data.iterrows():
                    product = row['Product']
                    revenue = row['Revenue']
                    
                    # 해당 제품의 행 찾기
                    for item in tree.get_children():
                        if tree.item(item)['values'][0] == product:
                            values = list(tree.item(item)['values'])
                            if fy == 'FY20':
                                values[1] = f"{unit_price:,.2f}"
                            elif fy == 'FY21':
                                values[2] = f"{unit_price:,.2f}"
                            elif fy == 'FY22':
                                values[3] = f"{unit_price:,.2f}"
                            elif fy == 'FY23':
                                values[4] = f"{unit_price:,.2f}"
                            elif fy == 'FY24':
                                values[5] = f"{unit_price:,.2f}"
                            elif fy == 'FY25':
                                values[6] = f"{unit_price:,.2f}"
                            tree.item(item, values=values)
                            break
                    else:
                        # 새로운 제품 행 추가
                        values = [product] + ['0.00', '0.00', '0.00', '0.00', '0.00', '0.00', '0.00']
                        if fy == 'FY20':
                            values[1] = f"{revenue:,.2f}"
                        elif fy == 'FY21':
                            values[2] = f"{revenue:,.2f}"
                        elif fy == 'FY22':
                            values[3] = f"{revenue:,.2f}"
                        elif fy == 'FY23':
                            values[4] = f"{revenue:,.2f}"
                        elif fy == 'FY24':
                            values[5] = f"{revenue:,.2f}"
                        elif fy == 'FY25':
                            values[6] = f"{revenue:,.2f}"
                        tree.insert('', 'end', values=values)
            
            # 총합계 행 추가
            total_revenue = product_data['Revenue'].sum()
            tree.insert('', 'end', values=['총합계', '', '', '', '', '', '', 
                                         f"{total_revenue:,.2f}"], 
                       tags=('total',))
        
        return tree
    
    def create_product_chart(self, parent, data):
        """제품별 차트 생성"""
        try:
            # 차트 제목
            title_label = tk.Label(parent, text="제품별 실적 차트", font=("맑은 고딕", 12, "bold"), 
                                  bg='white', fg='#2c3e50')
            title_label.pack(anchor='w', pady=(0, 10))
            
            # matplotlib 차트 생성
            fig, ax = plt.subplots(figsize=(10, 6))
            
            if 'Product' in data.columns and 'Revenue' in data.columns:
                # 제품별 총 매출 집계
                product_revenue = data.groupby('Product')['Revenue'].sum().sort_values(ascending=False)
                
                if not product_revenue.empty:
                    # 상위 10개 제품만 표시
                    top_products = product_revenue.head(10)
                    
                    products = top_products.index
                    revenues = top_products.values
                    
                    # 바 차트 생성
                    bars = ax.bar(range(len(products)), revenues, color='skyblue', alpha=0.7)
                    
                    # x축 라벨 설정
                    ax.set_xticks(range(len(products)))
                    ax.set_xticklabels(products, rotation=45, ha='right')
                    
                    # 값 표시
                    for bar, revenue in zip(bars, revenues):
                        height = bar.get_height()
                        ax.text(bar.get_x() + bar.get_width()/2., height + height*0.01,
                               f'{revenue:,.0f}', ha='center', va='bottom', fontsize=8)
                    
                    ax.set_title('제품별 매출 (상위 10개)', fontsize=14, fontweight='bold')
                    ax.set_ylabel('Revenue ($)', fontsize=12)
                    ax.set_xlabel('Product', fontsize=12)
                    ax.grid(True, alpha=0.3)
                else:
                    ax.text(0.5, 0.5, '데이터가 없습니다', ha='center', va='center', 
                           transform=ax.transAxes, fontsize=14)
            else:
                ax.text(0.5, 0.5, '데이터가 없습니다', ha='center', va='center', 
                       transform=ax.transAxes, fontsize=14)
            
            plt.tight_layout()
            
            # Tkinter에 차트 삽입
            canvas = FigureCanvasTkAgg(fig, parent)
            canvas.draw()
            canvas.get_tk_widget().pack(fill='both', expand=True)
            
        except Exception as e:
            print(f"제품별 차트 생성 실패: {e}")
            error_label = tk.Label(parent, text=f"차트 생성 중 오류가 발생했습니다.\n{e}",
                                  font=("맑은 고딕", 10), bg='white', fg='#e74c3c')
            error_label.pack(expand=True)
    
    def update_license_type_tab(self):
        """라이센스 타입별 탭 업데이트 - 이미지와 같은 형태"""
        try:
            # 기존 위젯 정리
            for widget in self.license_table_frame.winfo_children():
                widget.destroy()
            
            for widget in self.license_chart_frame.winfo_children():
                widget.destroy()
            
            # 필터링된 데이터
            filtered_df = self.get_filtered_data()
            
            if filtered_df.empty:
                # 데이터가 없을 때 메시지 표시
                no_data_label = tk.Label(self.license_table_frame, text="선택된 조건에 해당하는 데이터가 없습니다.",
                                        font=("맑은 고딕", 12), bg='white', fg='#7f8c8d')
                no_data_label.pack(expand=True)
                return
            
            # 라이센스 타입별 테이블 생성
            self.create_license_type_table(self.license_table_frame, filtered_df)
            
            # 라이센스 타입별 차트 생성
            self.create_license_type_chart(self.license_chart_frame, filtered_df)
            
        except Exception as e:
            print(f"라이센스 타입별 탭 업데이트 실패: {e}")
    
    def create_license_type_table(self, parent, data):
        """라이센스 타입별 테이블 생성 - 연도별로 확인 가능"""
        # 테이블 생성
        columns = ['라이센스 타입', 'FY20', 'FY21', 'FY22', 'FY23', 'FY24', 'FY25', '총합계']
        
        tree = ttk.Treeview(parent, columns=columns, show='headings', height=20)
        
        # 컬럼 설정
        for col in columns:
            tree.heading(col, text=col)
            if col == '라이센스 타입':
                tree.column(col, width=150, anchor='w')
            else:
                tree.column(col, width=100, anchor='e')
        
        # 스크롤바
        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        
        # 라이센스 타입별 데이터 집계 (Product Type 1 컬럼 사용, 필터링된 데이터)
        if 'Product Type 1' in data.columns and 'Revenue' in data.columns:
            # 라이센스 타입별로 데이터 집계
            license_data = data.groupby('Product Type 1').agg({
                'Revenue': 'sum'
            }).reset_index()
            
            # 연도별 데이터 집계
            for fy in ['FY20', 'FY21', 'FY22', 'FY23', 'FY24', 'FY25']:
                fy_data = data[data['FY'] == fy].groupby('Product Type 1').agg({
                    'Revenue': 'sum'
                }).reset_index()
                
                # 라이센스 타입별 데이터에 연도별 데이터 추가
                for _, row in fy_data.iterrows():
                    license_type = row['Product Type 1']
                    revenue = row['Revenue']
                    
                    # 해당 라이센스 타입의 행 찾기
                    for item in tree.get_children():
                        if tree.item(item)['values'][0] == license_type:
                            values = list(tree.item(item)['values'])
                            if fy == 'FY20':
                                values[1] = f"{revenue:,.2f}"
                            elif fy == 'FY21':
                                values[2] = f"{revenue:,.2f}"
                            elif fy == 'FY22':
                                values[3] = f"{revenue:,.2f}"
                            elif fy == 'FY23':
                                values[4] = f"{revenue:,.2f}"
                            elif fy == 'FY24':
                                values[5] = f"{revenue:,.2f}"
                            elif fy == 'FY25':
                                values[6] = f"{revenue:,.2f}"
                            tree.item(item, values=values)
                            break
                    else:
                        # 새로운 라이센스 타입 행 추가
                        values = [license_type] + ['0.00', '0.00', '0.00', '0.00', '0.00', '0.00', '0.00']
                        if fy == 'FY20':
                            values[1] = f"{revenue:,.2f}"
                        elif fy == 'FY21':
                            values[2] = f"{revenue:,.2f}"
                        elif fy == 'FY22':
                            values[3] = f"{revenue:,.2f}"
                        elif fy == 'FY23':
                            values[4] = f"{revenue:,.2f}"
                        elif fy == 'FY24':
                            values[5] = f"{revenue:,.2f}"
                        elif fy == 'FY25':
                            values[6] = f"{revenue:,.2f}"
                        tree.insert('', 'end', values=values)
            
            # 총합계 행 추가
            total_revenue = license_data['Revenue'].sum()
            tree.insert('', 'end', values=['총합계', '', '', '', '', '', '', 
                                         f"{total_revenue:,.2f}"], 
                       tags=('total',))
        
        return tree
    
    def create_license_type_chart(self, parent, data):
        """라이센스 타입별 차트 생성"""
        try:
            # 차트 제목
            title_label = tk.Label(parent, text="라이센스 타입별 실적 차트", font=("맑은 고딕", 12, "bold"), 
                                  bg='white', fg='#2c3e50')
            title_label.pack(anchor='w', pady=(0, 10))
            
            # matplotlib 차트 생성
            fig, ax = plt.subplots(figsize=(10, 6))
            
            if 'Product Type 1' in data.columns and 'Revenue' in data.columns:
                # 라이센스 타입별 총 매출 집계
                license_revenue = data.groupby('Product Type 1')['Revenue'].sum().sort_values(ascending=False)
                
                if not license_revenue.empty:
                    license_types = license_revenue.index
                    revenues = license_revenue.values
                    
                    # 바 차트 생성
                    bars = ax.bar(range(len(license_types)), revenues, color='lightcoral', alpha=0.7)
                    
                    # x축 라벨 설정
                    ax.set_xticks(range(len(license_types)))
                    ax.set_xticklabels(license_types, rotation=45, ha='right')
                    
                    # 값 표시
                    for bar, revenue in zip(bars, revenues):
                        height = bar.get_height()
                        ax.text(bar.get_x() + bar.get_width()/2., height + height*0.01,
                               f'{revenue:,.0f}', ha='center', va='bottom', fontsize=8)
                    
                    ax.set_title('라이센스 타입별 매출', fontsize=14, fontweight='bold')
                    ax.set_ylabel('Revenue ($)', fontsize=12)
                    ax.set_xlabel('License Type', fontsize=12)
                    ax.grid(True, alpha=0.3)
                else:
                    ax.text(0.5, 0.5, '데이터가 없습니다', ha='center', va='center', 
                           transform=ax.transAxes, fontsize=14)
            else:
                ax.text(0.5, 0.5, '데이터가 없습니다', ha='center', va='center', 
                       transform=ax.transAxes, fontsize=14)
            
            plt.tight_layout()
            
            # Tkinter에 차트 삽입
            canvas = FigureCanvasTkAgg(fig, parent)
            canvas.draw()
            canvas.get_tk_widget().pack(fill='both', expand=True)
            
        except Exception as e:
            print(f"라이센스 타입별 차트 생성 실패: {e}")
            error_label = tk.Label(parent, text=f"차트 생성 중 오류가 발생했습니다.\n{e}",
                                  font=("맑은 고딕", 10), bg='white', fg='#e74c3c')
            error_label.pack(expand=True)
    
    def update_customer_tab(self):
        """고객사별 탭 업데이트 - 테이블만 표시"""
        try:
            # 기존 위젯 정리
            for widget in self.customer_table_frame.winfo_children():
                widget.destroy()
            
            # 필터링된 데이터
            filtered_df = self.get_filtered_data()
            
            if filtered_df.empty:
                # 데이터가 없을 때 메시지 표시
                no_data_label = tk.Label(self.customer_table_frame, text="선택된 조건에 해당하는 데이터가 없습니다.",
                                        font=("맑은 고딕", 12), bg='white', fg='#7f8c8d')
                no_data_label.pack(expand=True)
                return
            
            # 고객사별 테이블 생성
            self.create_customer_table(self.customer_table_frame, filtered_df)
            
        except Exception as e:
            print(f"고객사별 탭 업데이트 실패: {e}")
    
    def create_customer_table(self, parent, data):
        """고객사별 테이블 생성 - 이미지와 같은 형태, 분기별 보기 기능"""
        # 보기 모드 선택 프레임
        view_frame = tk.Frame(parent, bg='white')
        view_frame.pack(fill='x', pady=(0, 10))
        
        # 보기 모드 라디오 버튼
        self.view_mode = tk.StringVar(value="연도별")
        tk.Label(view_frame, text="보기 모드:", font=("맑은 고딕", 10), bg='white').pack(side='left', padx=(0, 10))
        
        tk.Radiobutton(view_frame, text="연도별", variable=self.view_mode, value="연도별", 
                      command=self.on_view_mode_changed, bg='white').pack(side='left', padx=(0, 10))
        tk.Radiobutton(view_frame, text="분기별", variable=self.view_mode, value="분기별", 
                      command=self.on_view_mode_changed, bg='white').pack(side='left')
        
        # 테이블 생성
        if self.filter_mode.get() == "비교":
            # 비교 모드: 두 개의 비교 대상과 변화율
            compare1_year = self.compare1_year.get()
            compare1_quarter = self.compare1_quarter.get()
            compare2_year = self.compare2_year.get()
            compare2_quarter = self.compare2_quarter.get()
            
            if all([compare1_year, compare1_quarter, compare2_year, compare2_quarter]):
                columns = [
                    '고객사명', 
                    f'{compare1_year} {compare1_quarter}', 
                    f'{compare2_year} {compare2_quarter}', 
                    '변화율'
                ]
            else:
                columns = ['고객사명', '비교 대상 1', '비교 대상 2', '변화율']
        elif self.view_mode.get() == "연도별":
            columns = ['고객사명', 'FY20', 'FY21', 'FY22', 'FY23', 'FY24', 'FY25', '22/23 YoY', '23/24 YoY', '24/25 YoY']
        else:
            columns = ['고객사명', 'Q1', 'Q2', 'Q3', 'Q4', '총합계']
        
        tree = ttk.Treeview(parent, columns=columns, show='headings', height=20)
        
        # 컬럼 설정
        for col in columns:
            tree.heading(col, text=col)
            if col == '고객사명':
                tree.column(col, width=200, anchor='w')
            else:
                tree.column(col, width=100, anchor='e')
        
        # 스크롤바
        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        
        # 데이터 로드
        self.load_customer_data(tree, data)
        
        # 태그별 스타일 설정
        tree.tag_configure('customer', background='#e6f3ff')
        tree.tag_configure('total', background='#fff2e6')
        
        return tree
    
    def on_view_mode_changed(self):
        """보기 모드 변경 시 호출"""
        self.update_customer_tab()
    
    def load_customer_data(self, tree, data):
        """고객사 데이터 로드"""
        if 'Account' in data.columns and 'Revenue' in data.columns:
            if self.filter_mode.get() == "비교":
                self.load_compare_data(tree, data)
            elif self.view_mode.get() == "연도별":
                self.load_yearly_data(tree, data)
            else:
                self.load_quarterly_data(tree, data)
    
    def load_compare_data(self, tree, data):
        """비교 모드 데이터 로드"""
        if 'Compare_Group' not in data.columns:
            return
        
        # 비교 그룹별 데이터 집계
        compare_groups = data['Compare_Group'].unique()
        if len(compare_groups) != 2:
            return
        
        group1, group2 = compare_groups
        
        # 고객사별 비교 데이터 집계
        customer_compare = data.groupby(['Account', 'Compare_Group'])['Revenue'].sum().reset_index()
        
        # 고객사별로 데이터 정리
        customers = customer_compare['Account'].unique()
        
        for customer in customers:
            customer_data = customer_compare[customer_compare['Account'] == customer]
            
            # 각 그룹의 데이터
            group1_data = customer_data[customer_data['Compare_Group'] == group1]['Revenue'].sum()
            group2_data = customer_data[customer_data['Compare_Group'] == group2]['Revenue'].sum()
            
            # 변화율 계산
            if group1_data > 0:
                change_rate = ((group2_data - group1_data) / group1_data) * 100
            else:
                change_rate = 0
            
            # 행 추가
            values = [
                customer,
                f"{group1_data:,.0f}",
                f"{group2_data:,.0f}",
                f"{change_rate:+.1f}%" if change_rate != 0 else ""
            ]
            
            tree.insert('', 'end', values=values, tags=('customer',))
        
        # 총합계 계산
        total_group1 = customer_compare[customer_compare['Compare_Group'] == group1]['Revenue'].sum()
        total_group2 = customer_compare[customer_compare['Compare_Group'] == group2]['Revenue'].sum()
        
        if total_group1 > 0:
            total_change_rate = ((total_group2 - total_group1) / total_group1) * 100
        else:
            total_change_rate = 0
        
        total_values = [
            "총합계",
            f"{total_group1:,.0f}",
            f"{total_group2:,.0f}",
            f"{total_change_rate:+.1f}%" if total_change_rate != 0 else ""
        ]
        
        tree.insert('', 'end', values=total_values, tags=('total',))
    
    def load_yearly_data(self, tree, data):
        """연도별 데이터 로드"""
        # 고객사별 연도별 데이터 집계
        customer_yearly = data.groupby(['Account', 'FY'])['Revenue'].sum().reset_index()
        
        # 고객사별로 데이터 정리
        customers = customer_yearly['Account'].unique()
        
        for customer in customers:
            customer_data = customer_yearly[customer_yearly['Account'] == customer]
            
            # 연도별 데이터
            fy_data = {}
            for _, row in customer_data.iterrows():
                fy_data[row['FY']] = row['Revenue']
            
            # YoY 계산
            yoy_22_23 = self.calculate_yoy(fy_data.get('FY22', 0), fy_data.get('FY23', 0))
            yoy_23_24 = self.calculate_yoy(fy_data.get('FY23', 0), fy_data.get('FY24', 0))
            yoy_24_25 = self.calculate_yoy(fy_data.get('FY24', 0), fy_data.get('FY25', 0))
            
            # 행 추가
            values = [
                customer,
                f"{fy_data.get('FY20', 0):,.0f}",
                f"{fy_data.get('FY21', 0):,.0f}",
                f"{fy_data.get('FY22', 0):,.0f}",
                f"{fy_data.get('FY23', 0):,.0f}",
                f"{fy_data.get('FY24', 0):,.0f}",
                f"{fy_data.get('FY25', 0):,.0f}",
                f"{yoy_22_23:.1f}%" if yoy_22_23 != 0 else "",
                f"{yoy_23_24:.1f}%" if yoy_23_24 != 0 else "",
                f"{yoy_24_25:.1f}%" if yoy_24_25 != 0 else ""
            ]
            
            tree.insert('', 'end', values=values, tags=('customer',))
        
        # 총합계 계산
        total_fy20 = customer_yearly[customer_yearly['FY'] == 'FY20']['Revenue'].sum()
        total_fy21 = customer_yearly[customer_yearly['FY'] == 'FY21']['Revenue'].sum()
        total_fy22 = customer_yearly[customer_yearly['FY'] == 'FY22']['Revenue'].sum()
        total_fy23 = customer_yearly[customer_yearly['FY'] == 'FY23']['Revenue'].sum()
        total_fy24 = customer_yearly[customer_yearly['FY'] == 'FY24']['Revenue'].sum()
        total_fy25 = customer_yearly[customer_yearly['FY'] == 'FY25']['Revenue'].sum()
        
        total_yoy_22_23 = self.calculate_yoy(total_fy22, total_fy23)
        total_yoy_23_24 = self.calculate_yoy(total_fy23, total_fy24)
        total_yoy_24_25 = self.calculate_yoy(total_fy24, total_fy25)
        
        total_values = [
            "총합계",
            f"{total_fy20:,.0f}",
            f"{total_fy21:,.0f}",
            f"{total_fy22:,.0f}",
            f"{total_fy23:,.0f}",
            f"{total_fy24:,.0f}",
            f"{total_fy25:,.0f}",
            f"{total_yoy_22_23:.1f}%" if total_yoy_22_23 != 0 else "",
            f"{total_yoy_23_24:.1f}%" if total_yoy_23_24 != 0 else "",
            f"{total_yoy_24_25:.1f}%" if total_yoy_24_25 != 0 else ""
        ]
        
        tree.insert('', 'end', values=total_values, tags=('total',))
    
    def load_quarterly_data(self, tree, data):
        """분기별 데이터 로드"""
        # 고객사별 분기별 데이터 집계
        customer_quarterly = data.groupby(['Account', 'Quarter'])['Revenue'].sum().reset_index()
        
        # 고객사별로 데이터 정리
        customers = customer_quarterly['Account'].unique()
        
        for customer in customers:
            customer_data = customer_quarterly[customer_quarterly['Account'] == customer]
            
            # 분기별 데이터
            q1_total = customer_data[customer_data['Quarter'] == 'Q1']['Revenue'].sum()
            q2_total = customer_data[customer_data['Quarter'] == 'Q2']['Revenue'].sum()
            q3_total = customer_data[customer_data['Quarter'] == 'Q3']['Revenue'].sum()
            q4_total = customer_data[customer_data['Quarter'] == 'Q4']['Revenue'].sum()
            total = q1_total + q2_total + q3_total + q4_total
            
            # 행 추가
            values = [
                customer,
                f"{q1_total:,.0f}",
                f"{q2_total:,.0f}",
                f"{q3_total:,.0f}",
                f"{q4_total:,.0f}",
                f"{total:,.0f}"
            ]
            
            tree.insert('', 'end', values=values, tags=('customer',))
        
        # 총합계 계산
        total_q1 = customer_quarterly[customer_quarterly['Quarter'] == 'Q1']['Revenue'].sum()
        total_q2 = customer_quarterly[customer_quarterly['Quarter'] == 'Q2']['Revenue'].sum()
        total_q3 = customer_quarterly[customer_quarterly['Quarter'] == 'Q3']['Revenue'].sum()
        total_q4 = customer_quarterly[customer_quarterly['Quarter'] == 'Q4']['Revenue'].sum()
        total_all = total_q1 + total_q2 + total_q3 + total_q4
        
        total_values = [
            "총합계",
            f"{total_q1:,.0f}",
            f"{total_q2:,.0f}",
            f"{total_q3:,.0f}",
            f"{total_q4:,.0f}",
            f"{total_all:,.0f}"
        ]
        
        tree.insert('', 'end', values=total_values, tags=('total',))
    
    def calculate_yoy(self, prev_year, current_year):
        """전년 대비 성장률 계산"""
        if prev_year > 0:
            return ((current_year - prev_year) / prev_year) * 100
        return 0
    
    def create_customer_chart(self, parent, data):
        """고객사별 차트 생성"""
        try:
            # 차트 제목
            title_label = tk.Label(parent, text="고객사별 실적 차트", font=("맑은 고딕", 12, "bold"), 
                                  bg='white', fg='#2c3e50')
            title_label.pack(anchor='w', pady=(0, 10))
            
            # matplotlib 차트 생성
            fig, ax = plt.subplots(figsize=(10, 6))
            
            if 'Account' in data.columns and 'Revenue' in data.columns:
                # 고객사별 총 매출 집계
                customer_revenue = data.groupby('Account')['Revenue'].sum().sort_values(ascending=False)
                
                if not customer_revenue.empty:
                    # 상위 15개 고객사만 표시
                    top_customers = customer_revenue.head(15)
                    
                    customers = top_customers.index
                    revenues = top_customers.values
                    
                    # 바 차트 생성
                    bars = ax.bar(range(len(customers)), revenues, color='lightgreen', alpha=0.7)
                    
                    # x축 라벨 설정
                    ax.set_xticks(range(len(customers)))
                    ax.set_xticklabels(customers, rotation=45, ha='right')
                    
                    # 값 표시
                    for bar, revenue in zip(bars, revenues):
                        height = bar.get_height()
                        ax.text(bar.get_x() + bar.get_width()/2., height + height*0.01,
                               f'{revenue:,.0f}', ha='center', va='bottom', fontsize=8)
                    
                    ax.set_title('고객사별 매출 (상위 15개)', fontsize=14, fontweight='bold')
                    ax.set_ylabel('Revenue ($)', fontsize=12)
                    ax.set_xlabel('Customer', fontsize=12)
                    ax.grid(True, alpha=0.3)
                else:
                    ax.text(0.5, 0.5, '데이터가 없습니다', ha='center', va='center', 
                           transform=ax.transAxes, fontsize=14)
            else:
                ax.text(0.5, 0.5, '데이터가 없습니다', ha='center', va='center', 
                       transform=ax.transAxes, fontsize=14)
            
            plt.tight_layout()
            
            # Tkinter에 차트 삽입
            canvas = FigureCanvasTkAgg(fig, parent)
            canvas.draw()
            canvas.get_tk_widget().pack(fill='both', expand=True)
            
        except Exception as e:
            print(f"고객사별 차트 생성 실패: {e}")
            error_label = tk.Label(parent, text=f"차트 생성 중 오류가 발생했습니다.\n{e}",
                                  font=("맑은 고딕", 10), bg='white', fg='#e74c3c')
            error_label.pack(expand=True)

def create_foundry_dashboard(parent_frame):
    """Foundry 대시보드 생성 함수"""
    return FoundryDashboard(parent_frame)

if __name__ == "__main__":
    # 테스트용 코드
    root = tk.Tk()
    root.title("Foundry 대시보드 테스트")
    root.geometry("1400x800")
    
    dashboard = create_foundry_dashboard(root)
    
    root.mainloop()

