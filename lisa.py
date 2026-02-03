
import tkinter as tk
from tkinter import ttk, messagebox
import sys
import os
import traceback

# Add current directory to path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# 로깅 설정
try:
    from integrations.lisa_logging import setup_logger
    logger = setup_logger('main')
except ImportError:
    import logging
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger('main')


# 설정 파일 로드
try:
    from config import KEY_FILE, SCOPES
except ImportError:
    KEY_FILE = 'resources/keys/service_account.json'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

class LoginWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("LISA Login")
        self.root.geometry("300x200")
        
        # 화면 중앙 배치
        self.center_window(300, 200)
        
        tk.Label(root, text="LISA v2.01", font=("Arial", 16, "bold")).pack(pady=20)
        
        frame = tk.Frame(root)
        frame.pack(pady=10)
        
        tk.Label(frame, text="사용자 이름:").pack(side='left', padx=5)
        self.entry = tk.Entry(frame)
        self.entry.insert(0, "김한경") # 기본값
        self.entry.pack(side='left', padx=5)
        self.entry.bind('<Return>', lambda e: self.login())
        self.entry.focus_set()
        
        tk.Button(root, text="로그인", command=self.login, width=20, bg="#007bff", fg="white").pack(pady=20)
        
    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        
    def login(self):
        username = self.entry.get()
        if not username:
            messagebox.showwarning("경고", "이름을 입력하세요")
            return
            
        # 데이터 로더 초기화 (로그인 시점)
        try:
            from core.data_loader import DataLoader
            self.data_loader = DataLoader(KEY_FILE, SCOPES)
            logger.info("DataLoader initialized")
        except Exception as e:
            messagebox.showerror("오류", f"데이터 로더 초기화 실패:\n{e}")
            logger.error(f"DataLoader init failed: {traceback.format_exc()}")
            return

        # 메인 앱 전환
        for widget in self.root.winfo_children():
            widget.destroy()
            
        MainApp(self.root, username, self.data_loader)

class MainApp:
    def __init__(self, root, username, data_loader):
        self.root = root
        self.root.title(f"LISA v2.01 - {username}")
        self.root.geometry("1400x900")
        
        # 화면 중앙 배치
        self.center_window(1400, 900)
        
        self.username = username
        self.data_loader = data_loader
        
        # 스타일 설정
        self.setup_style()
        
        # 탭 컨테이너
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True)
        
        # 각 탭 로드
        self.load_tabs()
        
    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def setup_style(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TNotebook", tabposition='n')
        style.configure("TNotebook.Tab", padding=[15, 5], font=('Arial', 10))

    def load_tabs(self):
        # 1. Renewal 관리
        self.add_tab("Renewal 관리", "modules.renewal.renewal_preview", "RenewalMailerGUI", has_dataloader=True)
        
        # 2. 고객사 관리
        self.add_tab("고객사 관리", "modules.customer.customer_info", "CustomerInfoGUI")
        
        # 3. 견적서
        self.add_tab("견적서", "modules.quote.quote_info", "QuoteInfoGUI", has_dataloader=True)
        
        # 4. 매입/매출
        self.add_tab("매입/매출", "modules.purchase.purchase_sales_info", "PurchaseSalesInfo")
        
        # 5. 대시보드 (Foundry)
        self.add_dashboard_tab("대시보드", "modules.dashboard.dashboard_foundry", "create_foundry_dashboard")
        
        # 6. 라이선스 인증서
        self.add_tab("라이선스 인증서", "certificates.license_certificate_info", "LicenseCertificateInfo") # 클래스명 추정

    def add_tab(self, title, module_path, class_name, has_dataloader=False):
        frame = tk.Frame(self.notebook)
        self.notebook.add(frame, text=title)
        
        try:
            # 동적 import
            module = __import__(module_path, fromlist=[class_name])
            cls = getattr(module, class_name)
            
            # 클래스 인스턴스 생성
            if has_dataloader:
                try:
                    cls(frame, username=self.username, data_loader=self.data_loader)
                except TypeError:
                    # data_loader 인자를 안 받는 경우 (구버전 호환)
                    cls(frame, username=self.username)
            else:
                cls(frame, username=self.username)
                
            logger.info(f"Loaded tab: {title}")
        except Exception as e:
            error_msg = f"{title} 로드 실패:\n{e}"
            tk.Label(frame, text=error_msg, fg="red").pack(expand=True)
            logger.error(f"Failed to load {title}: {traceback.format_exc()}")
            print(error_msg)

    def add_dashboard_tab(self, title, module_path, func_name):
        frame = tk.Frame(self.notebook)
        self.notebook.add(frame, text=title)
        
        try:
            module = __import__(module_path, fromlist=[func_name])
            func = getattr(module, func_name)
            func(frame)
            logger.info(f"Loaded dashboard: {title}")
        except Exception as e:
            error_msg = f"{title} 로드 실패:\n{e}"
            tk.Label(frame, text=error_msg, fg="red").pack(expand=True)
            logger.error(f"Failed to load {title}: {traceback.format_exc()}")

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = LoginWindow(root)
        root.mainloop()
    except Exception as e:
        with open("crash_log.txt", "w") as f:
            f.write(traceback.format_exc())
        messagebox.showerror("치명적 오류", f"앱 실행 중 오류가 발생했습니다.\n{e}\n\ncrash_log.txt를 확인하세요.")
