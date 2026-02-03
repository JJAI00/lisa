# expiry_alarm.py
"""
LISA 만료일 알람 팝업 모듈
LISA 시작 시 영업사원별 이번달/다음달 만료 예정 라이선스를 알림합니다.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime, timedelta
import json
import os
from integrations.lisa_logging import setup_logger

# 모듈 로거 설정
logger = setup_logger('expiry_alarm')

# 알람 설정 파일 경로
ALARM_CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'alarm_config.json')
ALARM_STATE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'alarm_state.json')


class ExpiryAlarmManager:
    """만료일 알람 관리자"""
    
    # 기본 설정
    DEFAULT_CONFIG = {
        "alarm_range_days": 60,  # 60일 전부터 알람
        "show_on_startup": True,
        "auto_show_delay_ms": 3000,  # 3초 후 표시
        "exclude_status": ["Drop", "100%"],  # 제외할 상태
        "include_salesperson": []  # 비어있으면 전체, 특정 영업사원만 지정 가능
    }
    
    def __init__(self, data_loader=None, username=None):
        """
        Args:
            data_loader: LISA 데이터 로더 객체
            username: 현재 로그인한 사용자 이름
        """
        self.data_loader = data_loader
        self.username = username or '테스트유저'
        self.config = self._load_config()
        self.expiry_data = None
        
    def _load_config(self) -> dict:
        """알람 설정 로드"""
        if os.path.exists(ALARM_CONFIG_FILE):
            try:
                with open(ALARM_CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # 기본값과 병합
                    return {**self.DEFAULT_CONFIG, **config}
            except Exception as e:
                logger.warning(f"알람 설정 로드 실패: {e}")
        return self.DEFAULT_CONFIG.copy()
    
    def _save_config(self):
        """알람 설정 저장"""
        try:
            with open(ALARM_CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.error(f"알람 설정 저장 실패: {e}")
    
    def _check_already_shown_today(self) -> bool:
        """오늘 이미 알람을 표시했는지 확인"""
        if os.path.exists(ALARM_STATE_FILE):
            try:
                with open(ALARM_STATE_FILE, 'r', encoding='utf-8') as f:
                    state = json.load(f)
                    last_shown = state.get('last_shown_date', '')
                    today = datetime.now().strftime('%Y-%m-%d')
                    return last_shown == today
            except Exception:
                pass
        return False
    
    def _mark_shown_today(self):
        """오늘 알람 표시됨으로 기록"""
        try:
            state = {'last_shown_date': datetime.now().strftime('%Y-%m-%d')}
            with open(ALARM_STATE_FILE, 'w', encoding='utf-8') as f:
                json.dump(state, f)
        except Exception as e:
            logger.error(f"알람 상태 저장 실패: {e}")
    
    def check_expiring_licenses(self) -> dict:
        """
        만료 예정 라이선스 확인
        
        Returns:
            dict: {'this_month': [...], 'next_month': [...], 'urgent': [...]}
        """
        if not self.data_loader:
            logger.warning("데이터 로더가 없습니다.")
            return {'this_month': [], 'next_month': [], 'urgent': []}
        
        try:
            # 리뉴얼 데이터 가져오기
            renewal_df = self.data_loader.get_renewal_data()
            
            if renewal_df is None or renewal_df.empty:
                logger.info("리뉴얼 데이터가 없습니다.")
                return {'this_month': [], 'next_month': [], 'urgent': []}
            
            # 만료일 컬럼 확인
            date_col = None
            for col in ['만료일', '만료 일자', 'Expiry Date', 'expiry_date']:
                if col in renewal_df.columns:
                    date_col = col
                    break
            
            if not date_col:
                logger.warning("만료일 컬럼을 찾을 수 없습니다.")
                return {'this_month': [], 'next_month': [], 'urgent': []}
            
            # 날짜 변환
            renewal_df['_만료일_dt'] = pd.to_datetime(renewal_df[date_col], errors='coerce')
            df = renewal_df[renewal_df['_만료일_dt'].notna()].copy()
            
            if df.empty:
                logger.info("유효한 만료일 데이터가 없습니다.")
                return {'this_month': [], 'next_month': [], 'urgent': []}
            
            # Status 필터링 (제외할 상태 제거)
            if 'Status' in df.columns:
                exclude_status = self.config.get('exclude_status', ['Drop', '100%'])
                df = df[~df['Status'].isin(exclude_status)]
            
            # 영업사원 필터링 (설정된 경우)
            if self.config.get('include_salesperson'):
                if '영업사원' in df.columns:
                    df = df[df['영업사원'].isin(self.config['include_salesperson'])]
            
            # 현재 날짜 기준
            today = datetime.now()
            this_month_start = today.replace(day=1)
            
            # 다음달 계산
            if today.month == 12:
                next_month_start = today.replace(year=today.year + 1, month=1, day=1)
                next_month_end = today.replace(year=today.year + 1, month=2, day=1) - timedelta(days=1)
            else:
                next_month_start = today.replace(month=today.month + 1, day=1)
                if today.month == 11:
                    next_month_end = today.replace(year=today.year + 1, month=1, day=1) - timedelta(days=1)
                else:
                    next_month_end = today.replace(month=today.month + 2, day=1) - timedelta(days=1)
            
            this_month_end = next_month_start - timedelta(days=1)
            
            # 이번주 (긴급)
            week_end = today + timedelta(days=7)
            
            # 분류
            urgent = df[(df['_만료일_dt'] >= today) & (df['_만료일_dt'] <= week_end)]
            this_month = df[(df['_만료일_dt'] >= today) & (df['_만료일_dt'] <= this_month_end)]
            next_month = df[(df['_만료일_dt'] >= next_month_start) & (df['_만료일_dt'] <= next_month_end)]
            
            def to_list(subset):
                result = []
                for _, row in subset.iterrows():
                    result.append({
                        '고객사명': row.get('고객사명', ''),
                        '제품': row.get('제품', row.get('소분류', '')),
                        '만료일': row['_만료일_dt'].strftime('%Y-%m-%d'),
                        '영업사원': row.get('영업사원', ''),
                        'Status': row.get('Status', '')
                    })
                return result
            
            self.expiry_data = {
                'urgent': to_list(urgent),
                'this_month': to_list(this_month),
                'next_month': to_list(next_month)
            }
            
            logger.info(f"만료 예정: 긴급 {len(self.expiry_data['urgent'])}건, "
                       f"이번달 {len(self.expiry_data['this_month'])}건, "
                       f"다음달 {len(self.expiry_data['next_month'])}건")
            
            return self.expiry_data
            
        except Exception as e:
            logger.error(f"만료 예정 라이선스 확인 오류: {e}")
            import traceback
            traceback.print_exc()
            return {'this_month': [], 'next_month': [], 'urgent': []}
    
    def should_show_alarm(self) -> bool:
        """알람을 표시해야 하는지 확인"""
        if not self.config.get('show_on_startup', True):
            return False
        
        if self._check_already_shown_today():
            logger.info("오늘 이미 알람이 표시되었습니다.")
            return False
        
        return True


class ExpiryAlarmPopup(tk.Toplevel):
    """만료일 알람 팝업 창"""
    
    def __init__(self, parent, alarm_manager: ExpiryAlarmManager, on_detail_click=None):
        """
        Args:
            parent: 부모 윈도우
            alarm_manager: ExpiryAlarmManager 인스턴스
            on_detail_click: 자세히보기 클릭 시 콜백 함수
        """
        super().__init__(parent)
        self.alarm_manager = alarm_manager
        self.on_detail_click = on_detail_click
        self.expiry_data = alarm_manager.expiry_data or {}
        
        self._setup_window()
        self._create_widgets()
        
        # 알람 표시됨으로 기록
        alarm_manager._mark_shown_today()
    
    def _setup_window(self):
        """창 설정"""
        self.title("🔔 만료일 알림")
        self.geometry("500x400")
        self.resizable(True, True)
        
        # 화면 중앙에 배치
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'+{x}+{y}')
        
        # 최상위 창으로 설정
        self.transient(self.master)
        self.grab_set()
        
        # 배경색 설정
        self.configure(bg="#F7F9FB")
    
    def _create_widgets(self):
        """위젯 생성"""
        # 헤더
        header_frame = tk.Frame(self, bg="#3B82F6", height=60)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)
        
        header_label = tk.Label(
            header_frame,
            text=f"🔔 {self.alarm_manager.username}님의 만료 예정 알림",
            font=("맑은 고딕", 14, "bold"),
            fg="white",
            bg="#3B82F6"
        )
        header_label.pack(side='left', padx=20, pady=15)
        
        # 컨텐츠 영역 (스크롤 가능)
        content_canvas = tk.Canvas(self, bg="#F7F9FB", highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=content_canvas.yview)
        content_frame = tk.Frame(content_canvas, bg="#F7F9FB")
        
        content_canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        content_canvas.pack(side="left", fill="both", expand=True)
        
        canvas_window = content_canvas.create_window((0, 0), window=content_frame, anchor="nw")
        
        def on_frame_configure(e):
            content_canvas.configure(scrollregion=content_canvas.bbox("all"))
        content_frame.bind("<Configure>", on_frame_configure)
        
        def on_canvas_configure(e):
            content_canvas.itemconfig(canvas_window, width=e.width)
        content_canvas.bind("<Configure>", on_canvas_configure)
        
        # 긴급 알림 (이번주 만료)
        urgent_items = self.expiry_data.get('urgent', [])
        if urgent_items:
            self._create_section(content_frame, "🔴 긴급: 이번주 만료", urgent_items, "#FEE2E2", "#DC2626")
        
        # 이번달 만료
        this_month_items = self.expiry_data.get('this_month', [])
        if this_month_items:
            self._create_section(content_frame, "🟠 이번달 만료", this_month_items, "#FEF3C7", "#D97706")
        
        # 다음달 만료
        next_month_items = self.expiry_data.get('next_month', [])
        if next_month_items:
            self._create_section(content_frame, "🔵 다음달 만료", next_month_items, "#DBEAFE", "#2563EB")
        
        # 항목이 없는 경우
        if not urgent_items and not this_month_items and not next_month_items:
            no_data_label = tk.Label(
                content_frame,
                text="✅ 만료 예정인 라이선스가 없습니다!",
                font=("맑은 고딕", 12),
                fg="#10B981",
                bg="#F7F9FB",
                pady=40
            )
            no_data_label.pack(fill='x')
        
        # 버튼 프레임
        button_frame = tk.Frame(self, bg="#F7F9FB", height=60)
        button_frame.pack(fill='x', side='bottom')
        button_frame.pack_propagate(False)
        
        # 확인 버튼
        confirm_btn = tk.Button(
            button_frame,
            text="확인",
            command=self.destroy,
            font=("맑은 고딕", 10, "bold"),
            bg="#3B82F6",
            fg="white",
            relief="flat",
            padx=20,
            pady=8,
            cursor="hand2"
        )
        confirm_btn.pack(side='right', padx=20, pady=15)
        confirm_btn.bind('<Enter>', lambda e: confirm_btn.configure(bg="#2563EB"))
        confirm_btn.bind('<Leave>', lambda e: confirm_btn.configure(bg="#3B82F6"))
        
        # 자세히보기 버튼
        if self.on_detail_click:
            detail_btn = tk.Button(
                button_frame,
                text="📋 Renewal 관리로 이동",
                command=self._on_detail,
                font=("맑은 고딕", 10),
                bg="#10B981",
                fg="white",
                relief="flat",
                padx=15,
                pady=8,
                cursor="hand2"
            )
            detail_btn.pack(side='right', padx=5, pady=15)
            detail_btn.bind('<Enter>', lambda e: detail_btn.configure(bg="#059669"))
            detail_btn.bind('<Leave>', lambda e: detail_btn.configure(bg="#10B981"))
        
        # 오늘 다시 보지 않기 체크박스
        self.skip_today_var = tk.BooleanVar(value=False)
        skip_checkbox = tk.Checkbutton(
            button_frame,
            text="오늘 다시 보지 않기",
            variable=self.skip_today_var,
            font=("맑은 고딕", 9),
            bg="#F7F9FB",
            activebackground="#F7F9FB"
        )
        skip_checkbox.pack(side='left', padx=20, pady=15)
    
    def _create_section(self, parent, title: str, items: list, bg_color: str, title_color: str):
        """섹션 생성"""
        section_frame = tk.Frame(parent, bg=bg_color, relief="solid", bd=1)
        section_frame.pack(fill='x', padx=15, pady=10)
        
        # 섹션 헤더
        header = tk.Label(
            section_frame,
            text=f"{title} ({len(items)}건)",
            font=("맑은 고딕", 11, "bold"),
            fg=title_color,
            bg=bg_color,
            anchor='w'
        )
        header.pack(fill='x', padx=10, pady=(10, 5))
        
        # 항목 목록 (최대 5개만 표시)
        display_items = items[:5]
        for item in display_items:
            item_text = f"• {item['고객사명']} - {item['제품']} ({item['만료일']})"
            if item.get('영업사원'):
                item_text += f" [{item['영업사원']}]"
            
            item_label = tk.Label(
                section_frame,
                text=item_text,
                font=("맑은 고딕", 9),
                fg="#374151",
                bg=bg_color,
                anchor='w'
            )
            item_label.pack(fill='x', padx=15, pady=1)
        
        # 더 있으면 표시
        if len(items) > 5:
            more_label = tk.Label(
                section_frame,
                text=f"  ... 외 {len(items) - 5}건 더 있음",
                font=("맑은 고딕", 9, "italic"),
                fg="#6B7280",
                bg=bg_color,
                anchor='w'
            )
            more_label.pack(fill='x', padx=15, pady=(1, 10))
        else:
            # 하단 여백
            tk.Frame(section_frame, bg=bg_color, height=10).pack(fill='x')
    
    def _on_detail(self):
        """자세히보기 클릭"""
        if self.on_detail_click:
            self.on_detail_click()
        self.destroy()


def show_expiry_alarm(parent, data_loader, username, on_detail_click=None, force=False):
    """
    만료일 알람 팝업을 표시합니다.
    
    Args:
        parent: 부모 윈도우
        data_loader: LISA 데이터 로더
        username: 현재 사용자 이름
        on_detail_click: 자세히보기 클릭 시 콜백
        force: True면 오늘 이미 표시했어도 강제로 표시
    
    Returns:
        bool: 알람이 표시되었으면 True
    """
    try:
        manager = ExpiryAlarmManager(data_loader, username)
        
        if not force and not manager.should_show_alarm():
            logger.info("알람 표시 조건이 충족되지 않았습니다.")
            return False
        
        # 만료 예정 라이선스 확인
        expiry_data = manager.check_expiring_licenses()
        
        # 표시할 항목이 있는지 확인
        total_items = (len(expiry_data.get('urgent', [])) + 
                      len(expiry_data.get('this_month', [])) + 
                      len(expiry_data.get('next_month', [])))
        
        if total_items == 0:
            logger.info("표시할 만료 예정 항목이 없습니다.")
            return False
        
        # 팝업 표시
        popup = ExpiryAlarmPopup(parent, manager, on_detail_click)
        logger.info("만료일 알람 팝업이 표시되었습니다.")
        return True
        
    except Exception as e:
        logger.error(f"만료일 알람 표시 오류: {e}")
        return False


# 테스트용 메인
if __name__ == '__main__':
    # 테스트용 더미 데이터 로더
    class DummyDataLoader:
        def get_renewal_data(self):
            # 테스트 데이터 생성
            today = datetime.now()
            data = {
                '고객사명': ['테스트고객A', '테스트고객B', '테스트고객C', '테스트고객D'],
                '제품': ['Nuke', 'Hiero', 'Cinema4D', 'After Effects'],
                '만료일': [
                    (today + timedelta(days=3)).strftime('%Y-%m-%d'),
                    (today + timedelta(days=15)).strftime('%Y-%m-%d'),
                    (today + timedelta(days=35)).strftime('%Y-%m-%d'),
                    (today + timedelta(days=45)).strftime('%Y-%m-%d'),
                ],
                '영업사원': ['김한경', '김한경', '이영희', '박철수'],
                'Status': ['80%', '60%', '40%', '80%']
            }
            return pd.DataFrame(data)
    
    root = tk.Tk()
    root.withdraw()  # 메인 윈도우 숨김
    
    show_expiry_alarm(root, DummyDataLoader(), '테스트유저', force=True)
    
    root.mainloop()
