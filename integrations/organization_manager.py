import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from core.sendlog import KEY_FILE, SCOPES
import tkinter as tk
from tkinter import ttk, messagebox

class OrganizationManager:
    """조직 정보 관리 클래스"""
    
    def __init__(self):
        self.organization_data = None
        self.hierarchy = {}
        self.departments = []
        self.teams = []
        self.salespeople = []
        self.load_organization_data()
    
    def load_organization_data(self):
        """로그인 워크시트에서 조직 정보 로드"""
        try:
            from renewal_gui import LOGIN_SHEET_ID, LOGIN_SHEET_NAME
            
            creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
            gc = gspread.authorize(creds)
            ws = gc.open_by_key(LOGIN_SHEET_ID).worksheet(LOGIN_SHEET_NAME)
            
            # 모든 데이터 가져오기
            all_data = ws.get_all_values()
            if len(all_data) < 2:
                print("조직 데이터가 없습니다.")
                return
            
            # DataFrame 생성
            headers = all_data[0]
            data = all_data[1:]
            self.organization_data = pd.DataFrame(data, columns=headers)
            
            # 컬럼명 정리
            self.organization_data.columns = self.organization_data.columns.str.strip()
            
            # 조직 계층 구조 생성
            self.build_hierarchy()
            
            print(f"조직 정보 로드 완료: {len(self.organization_data)}명")
            
        except Exception as e:
            print(f"조직 정보 로드 오류: {e}")
            self.organization_data = pd.DataFrame()
    
    def build_hierarchy(self):
        """조직 계층 구조 생성"""
        if self.organization_data.empty:
            return
        
        # 필요한 컬럼 확인
        required_columns = ['부서', '팀', '영업사원']
        available_columns = [col for col in required_columns if col in self.organization_data.columns]
        
        if not available_columns:
            print("조직 정보 컬럼을 찾을 수 없습니다.")
            return
        
        # 부서별로 그룹화
        if '부서' in self.organization_data.columns:
            self.departments = sorted(self.organization_data['부서'].dropna().unique().tolist())
        
        # 팀별로 그룹화
        if '팀' in self.organization_data.columns:
            self.teams = sorted(self.organization_data['팀'].dropna().unique().tolist())
        
        # 영업사원별로 그룹화
        if '영업사원' in self.organization_data.columns:
            self.salespeople = sorted(self.organization_data['영업사원'].dropna().unique().tolist())
        
        # 계층 구조 생성
        self.hierarchy = {
            '회사': '큐브렉스',
            '부서': {},
            '팀': {},
            '영업사원': {}
        }
        
        # 부서별 팀 매핑
        if '부서' in self.organization_data.columns and '팀' in self.organization_data.columns:
            for dept in self.departments:
                dept_teams = self.organization_data[
                    self.organization_data['부서'] == dept
                ]['팀'].dropna().unique().tolist()
                self.hierarchy['부서'][dept] = sorted(dept_teams)
        
        # 팀별 영업사원 매핑
        if '팀' in self.organization_data.columns and '영업사원' in self.organization_data.columns:
            for team in self.teams:
                team_salespeople = self.organization_data[
                    self.organization_data['팀'] == team
                ]['영업사원'].dropna().unique().tolist()
                self.hierarchy['팀'][team] = sorted(team_salespeople)
    
    def get_departments(self):
        """부서 목록 반환"""
        return self.departments
    
    def get_teams_by_department(self, department):
        """특정 부서의 팀 목록 반환"""
        return self.hierarchy.get('부서', {}).get(department, [])
    
    def get_salespeople_by_team(self, team):
        """특정 팀의 영업사원 목록 반환"""
        return self.hierarchy.get('팀', {}).get(team, [])
    
    def get_all_teams(self):
        """전체 팀 목록 반환"""
        return self.teams
    
    def get_all_salespeople(self):
        """전체 영업사원 목록 반환"""
        return self.salespeople
    
    def get_salespeople_by_department(self, department):
        """특정 부서의 모든 영업사원 반환"""
        if self.organization_data.empty or '부서' not in self.organization_data.columns:
            return []
        
        dept_salespeople = self.organization_data[
            self.organization_data['부서'] == department
        ]['영업사원'].dropna().unique().tolist()
        return sorted(dept_salespeople)
    
    def get_user_info(self, user_id):
        """사용자 정보 반환"""
        if self.organization_data.empty or '아이디' not in self.organization_data.columns:
            return None
        
        user_row = self.organization_data[self.organization_data['아이디'] == user_id]
        if not user_row.empty:
            return user_row.iloc[0].to_dict()
        return None
    
    def get_user_permissions(self, user_id):
        """사용자 권한 정보 반환"""
        user_info = self.get_user_info(user_id)
        if not user_info:
            return {
                'level': 'none',
                'departments': [],
                'teams': [],
                'salespeople': []
            }
        
        # 사용자의 부서, 팀 정보
        user_dept = user_info.get('부서', '')
        user_team = user_info.get('팀', '')
        user_salesperson = user_info.get('영업사원', '')
        
        # 권한 레벨 결정 (관리자, 부서장, 팀장, 영업사원)
        if user_info.get('권한', '') == '관리자':
            level = 'admin'
            departments = self.departments
            teams = self.teams
            salespeople = self.salespeople
        elif user_info.get('권한', '') == '부서장':
            level = 'department_head'
            departments = [user_dept] if user_dept else []
            teams = self.get_teams_by_department(user_dept)
            salespeople = self.get_salespeople_by_department(user_dept)
        elif user_info.get('권한', '') == '팀장':
            level = 'team_leader'
            departments = [user_dept] if user_dept else []
            teams = [user_team] if user_team else []
            salespeople = self.get_salespeople_by_team(user_team)
        else:
            level = 'salesperson'
            departments = [user_dept] if user_dept else []
            teams = [user_team] if user_team else []
            salespeople = [user_salesperson] if user_salesperson else []
        
        return {
            'level': level,
            'departments': departments,
            'teams': teams,
            'salespeople': salespeople,
            'user_dept': user_dept,
            'user_team': user_team,
            'user_salesperson': user_salesperson
        }

class OrganizationFilterWidget:
    """조직 필터 위젯"""
    
    def __init__(self, parent, org_manager, username=None):
        self.parent = parent
        self.org_manager = org_manager
        self.username = username
        self.user_permissions = None
        
        if username:
            self.user_permissions = org_manager.get_user_permissions(username)
        
        self.filter_vars = {
            'department': tk.StringVar(value="전체"),
            'team': tk.StringVar(value="전체"),
            'salesperson': tk.StringVar(value="전체")
        }
        
        self.create_widgets()
        self.setup_callbacks()
    
    def create_widgets(self):
        """필터 위젯 생성"""
        # 필터 프레임
        filter_frame = tk.Frame(self.parent, bg="white", relief="solid", bd=1)
        filter_frame.pack(fill='x', pady=(0, 10))
        
        # 필터 헤더
        filter_header = tk.Frame(filter_frame, bg="#F1F5F9", height=40)
        filter_header.pack(fill='x')
        filter_header.pack_propagate(False)
        
        tk.Label(filter_header, text="조직 필터", font=("맑은 고딕", 12, "bold"), 
                fg="#374151", bg="#F1F5F9").pack(side='left', padx=15, pady=10)
        
        # 필터 내용
        filter_content = tk.Frame(filter_frame, bg="white")
        filter_content.pack(fill='x', padx=15, pady=15)
        
        # 부서 선택
        tk.Label(filter_content, text="부서:", font=("맑은 고딕", 10), 
                fg="#374151", bg="white").grid(row=0, column=0, padx=(0, 5), pady=5)
        
        self.dept_combo = ttk.Combobox(filter_content, textvariable=self.filter_vars['department'], 
                                      width=15, font=("맑은 고딕", 10))
        self.dept_combo.grid(row=0, column=1, padx=(0, 15), pady=5)
        
        # 팀 선택
        tk.Label(filter_content, text="팀:", font=("맑은 고딕", 10), 
                fg="#374151", bg="white").grid(row=0, column=2, padx=(0, 5), pady=5)
        
        self.team_combo = ttk.Combobox(filter_content, textvariable=self.filter_vars['team'], 
                                      width=15, font=("맑은 고딕", 10))
        self.team_combo.grid(row=0, column=3, padx=(0, 15), pady=5)
        
        # 영업사원 선택
        tk.Label(filter_content, text="영업사원:", font=("맑은 고딕", 10), 
                fg="#374151", bg="white").grid(row=0, column=4, padx=(0, 5), pady=5)
        
        self.salesperson_combo = ttk.Combobox(filter_content, textvariable=self.filter_vars['salesperson'], 
                                             width=15, font=("맑은 고딕", 10))
        self.salesperson_combo.grid(row=0, column=5, padx=(0, 15), pady=5)
        
        # 새로고침 버튼
        refresh_btn = tk.Button(filter_content, text="🔄 새로고침", command=self.refresh_filters,
                               font=("맑은 고딕", 10, "bold"), bg="#3B82F6", fg="white",
                               relief="flat", padx=15, pady=5)
        refresh_btn.grid(row=0, column=6, pady=5)
        
        # 초기 필터 값 설정
        self.update_filter_options()
    
    def setup_callbacks(self):
        """콜백 함수 설정"""
        self.dept_combo.bind('<<ComboboxSelected>>', self.on_department_changed)
        self.team_combo.bind('<<ComboboxSelected>>', self.on_team_changed)
    
    def update_filter_options(self):
        """필터 옵션 업데이트"""
        # 사용자 권한에 따른 필터 옵션 설정
        if self.user_permissions:
            dept_options = ["전체"] + self.user_permissions['departments']
            team_options = ["전체"] + self.user_permissions['teams']
            salesperson_options = ["전체"] + self.user_permissions['salespeople']
        else:
            dept_options = ["전체"] + self.org_manager.get_departments()
            team_options = ["전체"] + self.org_manager.get_all_teams()
            salesperson_options = ["전체"] + self.org_manager.get_all_salespeople()
        
        self.dept_combo['values'] = dept_options
        self.team_combo['values'] = team_options
        self.salesperson_combo['values'] = salesperson_options
    
    def on_department_changed(self, event=None):
        """부서 변경 시 팀 옵션 업데이트"""
        selected_dept = self.filter_vars['department'].get()
        
        if selected_dept == "전체":
            # 전체 선택 시 모든 팀 표시
            if self.user_permissions:
                team_options = ["전체"] + self.user_permissions['teams']
            else:
                team_options = ["전체"] + self.org_manager.get_all_teams()
        else:
            # 특정 부서 선택 시 해당 부서의 팀만 표시
            team_options = ["전체"] + self.org_manager.get_teams_by_department(selected_dept)
        
        self.team_combo['values'] = team_options
        self.filter_vars['team'].set("전체")
        self.filter_vars['salesperson'].set("전체")
        
        # 팀 변경 이벤트도 트리거
        self.on_team_changed()
    
    def on_team_changed(self, event=None):
        """팀 변경 시 영업사원 옵션 업데이트"""
        selected_team = self.filter_vars['team'].get()
        
        if selected_team == "전체":
            # 전체 선택 시 부서 내 모든 영업사원 표시
            selected_dept = self.filter_vars['department'].get()
            if selected_dept == "전체":
                if self.user_permissions:
                    salesperson_options = ["전체"] + self.user_permissions['salespeople']
                else:
                    salesperson_options = ["전체"] + self.org_manager.get_all_salespeople()
            else:
                salesperson_options = ["전체"] + self.org_manager.get_salespeople_by_department(selected_dept)
        else:
            # 특정 팀 선택 시 해당 팀의 영업사원만 표시
            salesperson_options = ["전체"] + self.org_manager.get_salespeople_by_team(selected_team)
        
        self.salesperson_combo['values'] = salesperson_options
        self.filter_vars['salesperson'].set("전체")
    
    def refresh_filters(self):
        """필터 새로고침"""
        # 조직 데이터 다시 로드
        self.org_manager.load_organization_data()
        
        # 사용자 권한 다시 확인
        if self.username:
            self.user_permissions = self.org_manager.get_user_permissions(self.username)
        
        # 필터 옵션 업데이트
        self.update_filter_options()
        
        # 현재 선택된 값들 유지 (가능한 경우)
        current_dept = self.filter_vars['department'].get()
        current_team = self.filter_vars['team'].get()
        current_salesperson = self.filter_vars['salesperson'].get()
        
        # 옵션 업데이트 후 이전 값 복원
        self.dept_combo.event_generate('<<ComboboxSelected>>')
        
        # 이전 값이 여전히 유효한지 확인하고 설정
        if current_dept in self.dept_combo['values']:
            self.filter_vars['department'].set(current_dept)
        if current_team in self.team_combo['values']:
            self.filter_vars['team'].set(current_team)
        if current_salesperson in self.salesperson_combo['values']:
            self.filter_vars['salesperson'].set(current_salesperson)
    
    def get_selected_filters(self):
        """선택된 필터 값 반환"""
        return {
            'department': self.filter_vars['department'].get(),
            'team': self.filter_vars['team'].get(),
            'salesperson': self.filter_vars['salesperson'].get()
        }
    
    def apply_filters_to_data(self, df):
        """데이터에 필터 적용"""
        if df.empty:
            return df
        
        filtered_df = df.copy()
        filters = self.get_selected_filters()
        
        # 부서 필터
        if filters['department'] != "전체" and '부서' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['부서'] == filters['department']]
        
        # 팀 필터
        if filters['team'] != "전체" and '팀' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['팀'] == filters['team']]
        
        # 영업사원 필터
        if filters['salesperson'] != "전체" and '영업사원' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['영업사원'] == filters['salesperson']]
        
        return filtered_df

# 전역 조직 관리자 인스턴스
_org_manager = None

def get_organization_manager():
    """전역 조직 관리자 인스턴스 반환"""
    global _org_manager
    if _org_manager is None:
        _org_manager = OrganizationManager()
    return _org_manager 