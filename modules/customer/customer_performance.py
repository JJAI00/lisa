import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.font_manager as fm
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='matplotlib')
import numpy as np

class CustomerPerformanceDialog:
    """고객사별 연도/분기별 실적 다이얼로그"""
    
    def __init__(self, parent, customer_name, data_loader=None):
        self.parent = parent
        self.customer_name = customer_name
        self.data_loader = data_loader
        
        # 창 생성
        self.window = tk.Toplevel(parent)
        self.window.title(f"{customer_name} - 연도/분기별 실적")
        self.window.geometry("1200x800")
        self.window.configure(bg="#F7F9FB")
        
        # 창을 부모 창 중앙에 위치
        self.window.transient(parent)
        self.window.grab_set()
        
        # 데이터 로드
        self.load_customer_data()
        
        # UI 생성
        self.create_widgets()
        
        # 데이터 표시
        self.display_performance_data()
        
    def load_customer_data(self):
        """고객사 데이터 로드"""
        try:
            if self.data_loader:
                self.renewals_df = self.data_loader.get_renewal_data()
            else:
                # 기존 방식으로 데이터 로드
                from core.data_loader import get_global_auth, safe_api_call
                gc = get_global_auth()
                
                def load_renewal_data():
                    sh = gc.open_by_key('17rfmG5DEOj1CC-iA6SEVIdsS8zP-UC7lGBT6w2LwBvQ')
                    ws = sh.worksheet('Sheet1')
                    return ws.get_all_records()
                
                all_data = safe_api_call(load_renewal_data)
                if all_data:
                    self.renewals_df = pd.DataFrame(all_data)
                    self.renewals_df.columns = self.renewals_df.columns.astype(str).str.strip()
                else:
                    self.renewals_df = pd.DataFrame()
            
            # 해당 고객사의 데이터만 필터링
            if not self.renewals_df.empty and '고객사명' in self.renewals_df.columns:
                self.customer_df = self.renewals_df[
                    self.renewals_df['고객사명'].str.contains(self.customer_name, na=False, case=False, regex=False)
                ].copy()
                
                # 날짜 컬럼 변환
                if '계산서 발행일' in self.customer_df.columns:
                    self.customer_df['계산서_날짜'] = pd.to_datetime(self.customer_df['계산서 발행일'], errors='coerce')
                
                # 금액 컬럼 숫자로 변환
                amount_cols = ['판매 합계', '원가계', '이익액']
                for col in amount_cols:
                    if col in self.customer_df.columns:
                        self.customer_df[col] = pd.to_numeric(self.customer_df[col], errors='coerce')
                
                print(f"고객사 '{self.customer_name}' 데이터 로드 완료: {len(self.customer_df)}행")
            else:
                self.customer_df = pd.DataFrame()
                
        except Exception as e:
            print(f"고객사 데이터 로드 실패: {e}")
            self.customer_df = pd.DataFrame()
    
    def create_widgets(self):
        """UI 위젯 생성"""
        # 헤더
        header_frame = tk.Frame(self.window, bg="#F7F9FB", height=60)
        header_frame.pack(fill='x', pady=(0, 20))
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(header_frame, text=f"📊 {self.customer_name} - 연도/분기별 실적", 
                              font=("맑은 고딕", 16, "bold"), 
                              fg="#1F2937", bg="#F7F9FB")
        title_label.pack(side='left', padx=20, pady=15)
        
        # 탭 컨트롤
        self.notebook = ttk.Notebook(self.window)
        self.notebook.pack(fill='both', expand=True, padx=20, pady=(0, 20))
        
        # 연도별 탭
        self.year_frame = tk.Frame(self.notebook, bg="#F7F9FB")
        self.notebook.add(self.year_frame, text="연도별 실적")
        
        # 분기별 탭
        self.quarter_frame = tk.Frame(self.notebook, bg="#F7F9FB")
        self.notebook.add(self.quarter_frame, text="분기별 실적")
        
        # 제품별 탭
        self.product_frame = tk.Frame(self.notebook, bg="#F7F9FB")
        self.notebook.add(self.product_frame, text="제품별 실적")
        
        # 연도별 실적 테이블
        self.create_year_table()
        
        # 분기별 실적 테이블
        self.create_quarter_table()
        
        # 제품별 실적 차트
        self.create_product_chart()
    
    def create_year_table(self):
        """연도별 실적 테이블 생성"""
        # 테이블 프레임
        table_frame = tk.Frame(self.year_frame, bg="white", relief="solid", bd=1)
        table_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # 헤더
        header_label = tk.Label(table_frame, text="연도별 실적 요약", 
                               font=("맑은 고딕", 12, "bold"), 
                               fg="#374151", bg="white")
        header_label.pack(pady=10)
        
        # 트리뷰
        columns = ('연도', '거래건수', '총 매출', '총 원가', '총 이익', '이익률')
        self.year_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=10)
        
        # 스크롤바
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.year_tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.year_tree.xview)
        self.year_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # 컬럼 설정
        for col in columns:
            self.year_tree.heading(col, text=col, command=lambda c=col: self.sort_year_table(c))
            self.year_tree.column(col, width=120, anchor='center')
        
        # 배치
        self.year_tree.pack(side='left', fill='both', expand=True, padx=10, pady=10)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        
        # 더블클릭 이벤트
        self.year_tree.bind('<Double-1>', self.on_year_double_click)
    
    def create_quarter_table(self):
        """분기별 실적 테이블 생성"""
        # 테이블 프레임
        table_frame = tk.Frame(self.quarter_frame, bg="white", relief="solid", bd=1)
        table_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # 헤더
        header_label = tk.Label(table_frame, text="분기별 실적 요약", 
                               font=("맑은 고딕", 12, "bold"), 
                               fg="#374151", bg="white")
        header_label.pack(pady=10)
        
        # 트리뷰
        columns = ('연도', '분기', '거래건수', '총 매출', '총 원가', '총 이익', '이익률')
        self.quarter_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)
        
        # 스크롤바
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.quarter_tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.quarter_tree.xview)
        self.quarter_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # 컬럼 설정
        for col in columns:
            self.quarter_tree.heading(col, text=col, command=lambda c=col: self.sort_quarter_table(c))
            self.quarter_tree.column(col, width=100, anchor='center')
        
        # 배치
        self.quarter_tree.pack(side='left', fill='both', expand=True, padx=10, pady=10)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        
        # 더블클릭 이벤트
        self.quarter_tree.bind('<Double-1>', self.on_quarter_double_click)
    
    def create_product_chart(self):
        """제품별 실적 차트 생성"""
        # 차트 프레임
        chart_frame = tk.Frame(self.product_frame, bg="white", relief="solid", bd=1)
        chart_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # 헤더
        header_label = tk.Label(chart_frame, text="제품별 매출 비율", 
                               font=("맑은 고딕", 12, "bold"), 
                               fg="#374151", bg="white")
        header_label.pack(pady=10)
        
        # 차트 영역
        self.chart_frame = tk.Frame(chart_frame, bg="white")
        self.chart_frame.pack(fill='both', expand=True, padx=20, pady=20)
    
    def display_performance_data(self):
        """실적 데이터 표시"""
        if self.customer_df.empty:
            messagebox.showinfo("알림", f"'{self.customer_name}'의 데이터가 없습니다.")
            return
        
        # 연도별 데이터 계산
        self.calculate_year_data()
        
        # 분기별 데이터 계산
        self.calculate_quarter_data()
        
        # 제품별 차트 생성
        self.create_product_pie_chart()
    
    def calculate_year_data(self):
        """연도별 데이터 계산"""
        if '계산서_날짜' not in self.customer_df.columns:
            return
        
        # 연도별 그룹핑
        self.customer_df['연도'] = self.customer_df['계산서_날짜'].dt.year
        
        year_data = self.customer_df.groupby('연도').agg({
            '판매 합계': 'sum',
            '원가계': 'sum',
            '이익액': 'sum'
        }).reset_index()
        
        year_data['거래건수'] = self.customer_df.groupby('연도').size().values
        year_data['이익률'] = (year_data['이익액'] / year_data['판매 합계'] * 100).round(2)
        
        # 트리뷰에 데이터 삽입
        self.year_tree.delete(*self.year_tree.get_children())
        for _, row in year_data.iterrows():
            values = [
                int(row['연도']),
                int(row['거래건수']),
                f"{int(row['판매 합계']):,}",
                f"{int(row['원가계']):,}",
                f"{int(row['이익액']):,}",
                f"{row['이익률']:.2f}%"
            ]
            self.year_tree.insert('', 'end', values=values)
        
        self.year_data = year_data
    
    def calculate_quarter_data(self):
        """분기별 데이터 계산"""
        if '계산서_날짜' not in self.customer_df.columns:
            return
        
        # 분기 계산
        self.customer_df['연도'] = self.customer_df['계산서_날짜'].dt.year
        self.customer_df['분기'] = self.customer_df['계산서_날짜'].dt.quarter
        
        quarter_data = self.customer_df.groupby(['연도', '분기']).agg({
            '판매 합계': 'sum',
            '원가계': 'sum',
            '이익액': 'sum'
        }).reset_index()
        
        quarter_data['거래건수'] = self.customer_df.groupby(['연도', '분기']).size().values
        quarter_data['이익률'] = (quarter_data['이익액'] / quarter_data['판매 합계'] * 100).round(2)
        
        # 트리뷰에 데이터 삽입
        self.quarter_tree.delete(*self.quarter_tree.get_children())
        for _, row in quarter_data.iterrows():
            values = [
                int(row['연도']),
                f"{int(row['분기'])}분기",
                int(row['거래건수']),
                f"{int(row['판매 합계']):,}",
                f"{int(row['원가계']):,}",
                f"{int(row['이익액']):,}",
                f"{row['이익률']:.2f}%"
            ]
            self.quarter_tree.insert('', 'end', values=values)
        
        self.quarter_data = quarter_data
    
    def create_product_pie_chart(self):
        """제품별 원형 차트 생성"""
        if self.customer_df.empty or '제품' not in self.customer_df.columns:
            return
        
        # 제품별 매출 집계
        product_sales = self.customer_df.groupby('제품')['판매 합계'].sum().sort_values(ascending=False)
        
        if product_sales.empty:
            return
        
        # 상위 10개 제품만 표시 (나머지는 '기타'로)
        if len(product_sales) > 10:
            top_products = product_sales.head(9)
            other_sales = product_sales.iloc[9:].sum()
            product_sales = pd.concat([top_products, pd.Series({'기타': other_sales})])
        
        # matplotlib 차트 생성
        try:
            # 한글 폰트 설정
            korean_fonts = ['나눔고딕', 'NanumGothic', '맑은 고딕', 'Malgun Gothic', 'Dotum', '돋움']
            font_found = False
            
            for font_name in korean_fonts:
                try:
                    font_prop = fm.FontProperties(family=font_name)
                    if font_prop.get_name() != 'DejaVu Sans':
                        plt.rcParams['font.family'] = font_name
                        plt.rcParams['font.sans-serif'] = [font_name] + [f for f in plt.rcParams['font.sans-serif'] if f != font_name]
                        font_found = True
                        break
                except:
                    continue
            
            if not font_found:
                plt.rcParams['font.family'] = 'DejaVu Sans'
            
            # 차트 생성
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 8))
            
            # 원형 차트
            colors = plt.cm.Set3(np.linspace(0, 1, len(product_sales)))
            wedges, texts, autotexts = ax1.pie(product_sales.values, 
                                              labels=product_sales.index, 
                                              autopct='%1.1f%%',
                                              colors=colors,
                                              startangle=90)
            
            ax1.set_title(f'{self.customer_name} - 제품별 매출 비율', fontsize=14, fontweight='bold')
            
            # 막대 차트
            product_sales.plot(kind='bar', ax=ax2, color='skyblue', edgecolor='black')
            ax2.set_title(f'{self.customer_name} - 제품별 매출액', fontsize=14, fontweight='bold')
            ax2.set_xlabel('제품')
            ax2.set_ylabel('매출액 (원)')
            ax2.tick_params(axis='x', rotation=45)
            
            # y축 천 단위 콤마
            ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{int(x):,}'))
            
            plt.tight_layout()
            
            # Tkinter에 차트 삽입
            canvas = FigureCanvasTkAgg(fig, self.chart_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill='both', expand=True)
            
        except Exception as e:
            print(f"제품별 차트 생성 실패: {e}")
            # 차트 생성 실패 시 간단한 텍스트 표시
            error_label = tk.Label(self.chart_frame, text=f"차트 생성 실패: {e}", 
                                  font=("맑은 고딕", 12), fg="red", bg="white")
            error_label.pack(pady=50)
    
    def sort_year_table(self, column):
        """연도별 테이블 정렬"""
        # 현재 정렬 상태 확인
        current_sort = getattr(self, 'year_sort_column', None)
        current_reverse = getattr(self, 'year_sort_reverse', False)
        
        # 정렬 방향 결정
        if current_sort == column:
            reverse = not current_reverse
        else:
            reverse = False
        
        # 데이터 정렬
        if hasattr(self, 'year_data') and not self.year_data.empty:
            if column == '연도':
                sorted_data = self.year_data.sort_values('연도', ascending=not reverse)
            elif column == '거래건수':
                sorted_data = self.year_data.sort_values('거래건수', ascending=not reverse)
            elif column == '총 매출':
                sorted_data = self.year_data.sort_values('판매 합계', ascending=not reverse)
            elif column == '총 원가':
                sorted_data = self.year_data.sort_values('원가계', ascending=not reverse)
            elif column == '총 이익':
                sorted_data = self.year_data.sort_values('이익액', ascending=not reverse)
            elif column == '이익률':
                sorted_data = self.year_data.sort_values('이익률', ascending=not reverse)
            else:
                return
            
            # 트리뷰 업데이트
            self.year_tree.delete(*self.year_tree.get_children())
            for _, row in sorted_data.iterrows():
                values = [
                    int(row['연도']),
                    int(row['거래건수']),
                    f"{int(row['판매 합계']):,}",
                    f"{int(row['원가계']):,}",
                    f"{int(row['이익액']):,}",
                    f"{row['이익률']:.2f}%"
                ]
                self.year_tree.insert('', 'end', values=values)
            
            # 정렬 상태 저장
            self.year_sort_column = column
            self.year_sort_reverse = reverse
    
    def sort_quarter_table(self, column):
        """분기별 테이블 정렬"""
        # 현재 정렬 상태 확인
        current_sort = getattr(self, 'quarter_sort_column', None)
        current_reverse = getattr(self, 'quarter_sort_reverse', False)
        
        # 정렬 방향 결정
        if current_sort == column:
            reverse = not current_reverse
        else:
            reverse = False
        
        # 데이터 정렬
        if hasattr(self, 'quarter_data') and not self.quarter_data.empty:
            if column == '연도':
                sorted_data = self.quarter_data.sort_values('연도', ascending=not reverse)
            elif column == '분기':
                sorted_data = self.quarter_data.sort_values(['연도', '분기'], ascending=not reverse)
            elif column == '거래건수':
                sorted_data = self.quarter_data.sort_values('거래건수', ascending=not reverse)
            elif column == '총 매출':
                sorted_data = self.quarter_data.sort_values('판매 합계', ascending=not reverse)
            elif column == '총 원가':
                sorted_data = self.quarter_data.sort_values('원가계', ascending=not reverse)
            elif column == '총 이익':
                sorted_data = self.quarter_data.sort_values('이익액', ascending=not reverse)
            elif column == '이익률':
                sorted_data = self.quarter_data.sort_values('이익률', ascending=not reverse)
            else:
                return
            
            # 트리뷰 업데이트
            self.quarter_tree.delete(*self.quarter_tree.get_children())
            for _, row in sorted_data.iterrows():
                values = [
                    int(row['연도']),
                    f"{int(row['분기'])}분기",
                    int(row['거래건수']),
                    f"{int(row['판매 합계']):,}",
                    f"{int(row['원가계']):,}",
                    f"{int(row['이익액']):,}",
                    f"{row['이익률']:.2f}%"
                ]
                self.quarter_tree.insert('', 'end', values=values)
            
            # 정렬 상태 저장
            self.quarter_sort_column = column
            self.quarter_sort_reverse = reverse
    
    def on_year_double_click(self, event):
        """연도별 테이블 더블클릭 이벤트"""
        item = self.year_tree.identify_row(event.y)
        if not item:
            return
        
        values = self.year_tree.item(item, 'values')
        if not values:
            return
        
        year = int(values[0])
        self.show_year_detail(year)
    
    def on_quarter_double_click(self, event):
        """분기별 테이블 더블클릭 이벤트"""
        item = self.quarter_tree.identify_row(event.y)
        if not item:
            return
        
        values = self.quarter_tree.item(item, 'values')
        if not values:
            return
        
        year = int(values[0])
        quarter = int(values[1].replace('분기', ''))
        self.show_quarter_detail(year, quarter)
    
    def show_year_detail(self, year):
        """연도별 상세 데이터 표시"""
        if self.customer_df.empty:
            return
        
        # 해당 연도 데이터 필터링
        year_data = self.customer_df[self.customer_df['계산서_날짜'].dt.year == year].copy()
        
        if year_data.empty:
            messagebox.showinfo("알림", f"{year}년 데이터가 없습니다.")
            return
        
        # 상세 데이터 창 생성
        self.create_detail_window(f"{self.customer_name} - {year}년 상세 데이터", year_data)
    
    def show_quarter_detail(self, year, quarter):
        """분기별 상세 데이터 표시"""
        if self.customer_df.empty:
            return
        
        # 해당 분기 데이터 필터링
        quarter_data = self.customer_df[
            (self.customer_df['계산서_날짜'].dt.year == year) & 
            (self.customer_df['계산서_날짜'].dt.quarter == quarter)
        ].copy()
        
        if quarter_data.empty:
            messagebox.showinfo("알림", f"{year}년 {quarter}분기 데이터가 없습니다.")
            return
        
        # 상세 데이터 창 생성
        self.create_detail_window(f"{self.customer_name} - {year}년 {quarter}분기 상세 데이터", quarter_data)
    
    def create_detail_window(self, title, data):
        """상세 데이터 창 생성"""
        detail_window = tk.Toplevel(self.window)
        detail_window.title(title)
        detail_window.geometry("1000x600")
        detail_window.configure(bg="#F7F9FB")
        
        # 헤더
        header_frame = tk.Frame(detail_window, bg="#F7F9FB", height=50)
        header_frame.pack(fill='x', pady=(0, 10))
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(header_frame, text=title, 
                              font=("맑은 고딕", 14, "bold"), 
                              fg="#1F2937", bg="#F7F9FB")
        title_label.pack(side='left', padx=20, pady=10)
        
        # 테이블 프레임
        table_frame = tk.Frame(detail_window, bg="white", relief="solid", bd=1)
        table_frame.pack(fill='both', expand=True, padx=20, pady=(0, 20))
        
        # 컬럼 선택
        columns = ['계산서 발행일', '거래처명', '제품', '수량', '판매 단가', '판매 합계', '원가', '원가계', '이익액', '이익율']
        available_columns = [col for col in columns if col in data.columns]
        
        # 트리뷰
        tree = ttk.Treeview(table_frame, columns=available_columns, show='headings', height=20)
        
        # 스크롤바
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # 컬럼 설정
        for col in available_columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor='center')
        
        # 데이터 삽입
        for _, row in data.iterrows():
            values = []
            for col in available_columns:
                val = row.get(col, '')
                if col in ['판매 단가', '판매 합계', '원가', '원가계', '이익액']:
                    try:
                        val = f"{int(val):,}"
                    except:
                        val = str(val)
                elif col == '이익율':
                    try:
                        val = f"{float(val):.2f}%"
                    except:
                        val = str(val)
                values.append(val)
            tree.insert('', 'end', values=values)
        
        # 배치
        tree.pack(side='left', fill='both', expand=True, padx=10, pady=10)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
