import tkinter as tk
from tkinter import ttk
import threading
import time

class LoadingIndicator:
    """로딩 상태 표시 UI 컴포넌트"""
    
    def __init__(self, parent, title="데이터 로딩 중"):
        self.parent = parent
        self.title = title
        self.window = None
        self.progress_bar = None
        self.status_label = None
        self.progress_label = None
        self.is_visible = False
        
    def show(self, message="데이터를 로딩하고 있습니다..."):
        """로딩 창 표시"""
        if self.is_visible:
            return
            
        self.window = tk.Toplevel(self.parent)
        self.window.title(self.title)
        self.window.geometry("400x150")
        self.window.resizable(False, False)
        self.window.configure(bg="#F7F9FB")
        
        # 창을 화면 중앙에 배치
        self.window.transient(self.parent)
        self.window.grab_set()
        
        # 메인 프레임
        main_frame = tk.Frame(self.window, bg="#F7F9FB")
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # 로딩 아이콘 (텍스트 기반)
        icon_label = tk.Label(main_frame, text="⏳", font=("맑은 고딕", 24), 
                             bg="#F7F9FB", fg="#3B82F6")
        icon_label.pack(pady=(0, 10))
        
        # 상태 메시지
        self.status_label = tk.Label(main_frame, text=message, 
                                    font=("맑은 고딕", 11), 
                                    bg="#F7F9FB", fg="#374151")
        self.status_label.pack(pady=(0, 15))
        
        # 프로그레스 바
        self.progress_bar = ttk.Progressbar(main_frame, length=300, mode='determinate')
        self.progress_bar.pack(pady=(0, 5))
        
        # 진행률 텍스트
        self.progress_label = tk.Label(main_frame, text="0%", 
                                      font=("맑은 고딕", 10), 
                                      bg="#F7F9FB", fg="#6B7280")
        self.progress_label.pack()
        
        # 창을 화면 중앙에 배치
        self.center_window()
        
        self.is_visible = True
        
        # 강제 닫기 버튼 추가
        close_button = tk.Button(main_frame, text="X", font=("맑은 고딕", 10, "bold"),
                                bg="#ef4444", fg="white", relief="flat",
                                command=self.force_close)
        close_button.pack(pady=(10, 0))
        
        # 창이 닫히지 않도록 설정
        self.window.protocol("WM_DELETE_WINDOW", lambda: None)
        
    def update_progress(self, progress, status=""):
        """진행률 업데이트"""
        if not self.is_visible or not self.window:
            return
            
        try:
            # 진행률 제한 (0-100)
            progress = max(0, min(100, progress))
            
            # 프로그레스 바 업데이트
            if self.progress_bar and self.progress_bar.winfo_exists():
                self.progress_bar['value'] = progress
                
            # 진행률 텍스트 업데이트
            if self.progress_label and self.progress_label.winfo_exists():
                self.progress_label.config(text=f"{progress}%")
                
            # 상태 메시지 업데이트
            if self.status_label and self.status_label.winfo_exists() and status:
                self.status_label.config(text=status)
                
            # UI 업데이트
            if self.window and self.window.winfo_exists():
                self.window.update()
        except tk.TclError as e:
            # 위젯이 이미 파괴된 경우 무시
            print(f"로딩 인디케이터 업데이트 중 위젯 오류 (무시됨): {e}")
            self.is_visible = False
        except Exception as e:
            print(f"로딩 인디케이터 업데이트 오류: {e}")
            self.is_visible = False
        
    def hide(self):
        """로딩 창 숨기기"""
        if self.window and self.is_visible:
            self.window.destroy()
            self.window = None
            self.is_visible = False
            
    def force_close(self):
        """강제 닫기"""
        print("사용자가 로딩 창을 강제로 닫았습니다.")
        self.hide()
        
    def center_window(self):
        """창을 화면 중앙에 배치"""
        if not self.window:
            return
            
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f"{width}x{height}+{x}+{y}")

class BackgroundLoader:
    """백그라운드 로딩 관리자"""
    
    def __init__(self, parent, data_loader, on_complete=None):
        self.parent = parent
        self.data_loader = data_loader
        self.on_complete = on_complete
        self.loading_indicator = LoadingIndicator(parent)
        self.monitor_thread = None
        
    def start_loading(self):
        """백그라운드 로딩 시작"""
        # 로딩 인디케이터 표시
        self.loading_indicator.show("이번달 데이터를 먼저 로딩하고 있습니다...")
        
        # 이번달 데이터 로드
        current_month_data = self.data_loader.load_current_month_data()
        
        # 로딩 인디케이터 업데이트
        self.loading_indicator.update_progress(20, "이번달 데이터 로딩 완료")
        
        # 백그라운드에서 전체 데이터 로드 시작
        self.data_loader.load_full_data_async(callback=self.on_loading_complete)
        
        # 로딩 상태 모니터링 시작
        self.start_monitoring()
        
    def start_monitoring(self):
        """로딩 상태 모니터링"""
        def monitor():
            max_wait_time = 30  # 최대 30초 대기로 단축
            wait_count = 0
            
            while self.data_loader.is_loading and wait_count < max_wait_time * 10:  # 0.1초 간격으로 체크
                try:
                    status = self.data_loader.get_loading_status()
                    self.loading_indicator.update_progress(
                        status['progress'], 
                        status['status']
                    )
                except Exception as e:
                    print(f"로딩 상태 모니터링 오류: {e}")
                    break
                time.sleep(0.1)
                wait_count += 1
                
            # 로딩 완료 또는 타임아웃
            try:
                if wait_count >= max_wait_time * 10:
                    # 타임아웃 발생
                    print("로딩 타임아웃 발생 - 강제 완료 처리")
                    self.loading_indicator.update_progress(100, "로딩 완료 (타임아웃)")
                else:
                    self.loading_indicator.update_progress(100, "데이터 로딩 완료")
                
                time.sleep(0.5)  # 완료 메시지 잠시 표시
                self.loading_indicator.hide()
            except Exception as e:
                print(f"로딩 완료 처리 오류: {e}")
            
            # 완료 콜백 호출
            if self.on_complete:
                try:
                    self.on_complete()
                except Exception as e:
                    print(f"완료 콜백 실행 오류: {e}")
        
        self.monitor_thread = threading.Thread(target=monitor, daemon=True)
        self.monitor_thread.start()
        
    def stop_loading(self):
        """로딩 중단"""
        self.loading_indicator.hide()
        if self.monitor_thread and self.monitor_thread.is_alive():
            # 스레드는 daemon=True로 설정되어 있어서 자동으로 종료됨
            pass
    
    def on_loading_complete(self):
        """로딩 완료 시 호출되는 콜백"""
        # 로딩 완료 시 로딩 창 닫기
        self.loading_indicator.update_progress(100, "데이터 로딩 완료")
        time.sleep(0.5)  # 완료 메시지 잠시 표시
        self.loading_indicator.hide()
        
        # 완료 콜백 호출
        if self.on_complete:
            self.on_complete() 