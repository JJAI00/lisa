import tkinter as tk
from tkinter import ttk
from ui.column_settings import column_settings

class ResizableTreeview:
    """컬럼 크기 조절 기능이 있는 Treeview 믹스인"""
    
    def __init__(self, tree: ttk.Treeview, table_name: str):
        self.tree = tree
        self.table_name = table_name
        self.resize_column = None
        self.resize_start_x = 0
        self.resize_start_width = 0
        
        # 마우스 이벤트 바인딩
        self.tree.bind("<Button-1>", self.on_tree_click)
        self.tree.bind("<B1-Motion>", self.on_tree_drag)
        self.tree.bind("<ButtonRelease-1>", self.on_tree_release)
        
        # 컬럼 헤더 더블클릭으로 자동 크기 조절
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        
        # 우클릭 컨텍스트 메뉴
        self.tree.bind("<Button-3>", self.show_context_menu)
        
        # 저장된 설정 로드
        self.load_column_settings()
    
    def show_context_menu(self, event):
        """우클릭 컨텍스트 메뉴 표시"""
        region = self.tree.identify_region(event.x, event.y)
        
        if region == "heading":
            # 헤더 우클릭 시 컬럼 관련 메뉴
            self.show_column_context_menu(event)
        elif region == "cell":
            # 셀 우클릭 시 일반 메뉴
            self.show_cell_context_menu(event)
    
    def show_column_context_menu(self, event):
        """컬럼 헤더 우클릭 메뉴"""
        menu = tk.Menu(self.tree, tearoff=0)
        
        # 현재 컬럼 정보
        column = self.tree.identify_column(event.x)
        if column:
            try:
                col_name = self.tree['columns'][int(column.replace('#', '')) - 1]
                menu.add_command(label=f"'{col_name}' 컬럼 자동 크기 조절", 
                               command=lambda: self.auto_resize_column(column))
                menu.add_separator()
            except:
                pass
        
        menu.add_command(label="모든 컬럼 자동 크기 조절", 
                        command=self.auto_resize_all_columns)
        menu.add_command(label="컬럼 크기 기본값으로 리셋", 
                        command=self.reset_column_widths)
        menu.add_separator()
        menu.add_command(label="컬럼 설정 저장", 
                        command=self.save_column_settings)
        menu.add_command(label="컬럼 설정 불러오기", 
                        command=self.load_column_settings)
        
        # 메뉴 표시
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
    
    def show_cell_context_menu(self, event):
        """셀 우클릭 메뉴"""
        menu = tk.Menu(self.tree, tearoff=0)
        
        menu.add_command(label="모든 컬럼 자동 크기 조절", 
                        command=self.auto_resize_all_columns)
        menu.add_command(label="컬럼 크기 기본값으로 리셋", 
                        command=self.reset_column_widths)
        
        # 메뉴 표시
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
    
    def auto_resize_all_columns(self):
        """모든 컬럼 자동 크기 조절"""
        try:
            for col in self.tree['columns']:
                col_index = self.tree['columns'].index(col)
                column_id = f"#{col_index + 1}"
                self.auto_resize_column(column_id)
            print("모든 컬럼 자동 크기 조절 완료")
        except Exception as e:
            print(f"모든 컬럼 자동 크기 조절 오류: {e}")
    
    def load_column_settings(self):
        """저장된 컬럼 설정 로드"""
        try:
            # 컬럼 너비 설정 로드
            saved_widths = column_settings.get_column_widths(self.table_name)
            if saved_widths:
                for col, width in saved_widths.items():
                    try:
                        self.tree.column(col, width=width)
                    except:
                        pass  # 컬럼이 존재하지 않는 경우 무시
                print(f"{self.table_name} 컬럼 설정 로드 완료")
            
            # 컬럼 순서 설정 로드 (필요한 경우)
            saved_order = column_settings.get_column_order(self.table_name)
            if saved_order:
                # Treeview는 컬럼 순서 변경을 직접 지원하지 않으므로 
                # 필요한 경우 별도 구현 필요
                pass
                
        except Exception as e:
            print(f"컬럼 설정 로드 오류: {e}")
    
    def save_column_settings(self):
        """현재 컬럼 설정 저장"""
        try:
            # 현재 컬럼 너비 수집
            column_widths = {}
            for col in self.tree['columns']:
                try:
                    width = self.tree.column(col, 'width')
                    column_widths[col] = width
                except:
                    pass
            
            # 설정 저장
            column_settings.set_column_widths(self.table_name, column_widths)
            
            # 컬럼 순서도 저장 (필요한 경우)
            column_order = list(self.tree['columns'])
            column_settings.set_column_order(self.table_name, column_order)
            
            print(f"{self.table_name} 컬럼 설정 저장 완료")
            
        except Exception as e:
            print(f"컬럼 설정 저장 오류: {e}")
    
    def on_tree_click(self, event):
        """트리 클릭 이벤트"""
        region = self.tree.identify_region(event.x, event.y)
        if region == "separator":
            # 컬럼 구분선 클릭 시 리사이즈 모드 시작
            self.resize_column = self.tree.identify_column(event.x)
            self.resize_start_x = event.x
            try:
                self.resize_start_width = self.tree.column(self.resize_column, 'width')
            except:
                self.resize_start_width = 100
        else:
            self.resize_column = None
    
    def on_tree_drag(self, event):
        """트리 드래그 이벤트 (컬럼 크기 조절)"""
        if self.resize_column:
            try:
                # 드래그 거리 계산
                delta_x = event.x - self.resize_start_x
                new_width = max(20, self.resize_start_width + delta_x)  # 최소 20px
                
                # 컬럼 크기 업데이트
                self.tree.column(self.resize_column, width=new_width)
                
                # UI 업데이트
                self.tree.update_idletasks()
            except Exception as e:
                print(f"컬럼 크기 조절 오류: {e}")
    
    def on_tree_release(self, event):
        """트리 릴리즈 이벤트 (설정 저장)"""
        if self.resize_column:
            # 컬럼 크기 조절 완료 시 설정 저장
            self.save_column_settings()
            self.resize_column = None
    
    def on_tree_double_click(self, event):
        """트리 더블클릭 이벤트 (자동 크기 조절)"""
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            # 헤더 더블클릭 시 해당 컬럼 자동 크기 조절
            column = self.tree.identify_column(event.x)
            self.auto_resize_column(column)
    
    def auto_resize_column(self, column):
        """컬럼 자동 크기 조절"""
        try:
            # 컬럼의 모든 값 중 가장 긴 텍스트 찾기
            max_width = 0
            
            # 헤더 텍스트 길이 확인
            try:
                header_text = self.tree.heading(column, 'text')
                header_width = len(header_text) * 8  # 대략적인 픽셀 계산
                max_width = max(max_width, header_width)
            except:
                pass
            
            # 데이터 값들의 길이 확인
            for item in self.tree.get_children():
                try:
                    values = self.tree.item(item, 'values')
                    if values:
                        col_index = int(column.replace('#', '')) - 1
                        if col_index < len(values):
                            text = str(values[col_index])
                            text_width = len(text) * 8  # 대략적인 픽셀 계산
                            max_width = max(max_width, text_width)
                except:
                    pass
            
            # 여백 추가 (20px)
            new_width = max(50, max_width + 20)
            
            # 컬럼 크기 설정
            self.tree.column(column, width=new_width)
            
            # 설정 저장
            self.save_column_settings()
            
        except Exception as e:
            print(f"자동 크기 조절 오류: {e}")
    
    def reset_column_widths(self):
        """컬럼 크기를 기본값으로 리셋"""
        try:
            # 기본 너비 설정 (컬럼별로 다를 수 있음)
            default_widths = {
                '고객사명': 150,
                '거래처명': 150,
                '제품': 200,
                '만료일': 100,
                '계산서 발행일': 120,
                '판매 합계': 120,
                '이익액': 100,
                '영업사원': 100,
                '부서': 120,
                '팀': 120,
                '일자': 100,
                '금액': 120,
                '상태': 80,
                'Status': 80,
                '구분': 100,
                '연도': 80,
                '달성율': 80,
                '목표': 100,
                '실적': 100,
                '예상': 100
            }
            
            for col in self.tree['columns']:
                width = default_widths.get(col, 100)
                self.tree.column(col, width=width)
            
            # 설정 저장
            self.save_column_settings()
            
        except Exception as e:
            print(f"컬럼 크기 리셋 오류: {e}")
    
    def get_column_widths(self):
        """현재 컬럼 너비 반환"""
        widths = {}
        for col in self.tree['columns']:
            try:
                widths[col] = self.tree.column(col, 'width')
            except:
                widths[col] = 100
        return widths
