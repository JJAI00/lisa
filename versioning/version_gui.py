import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import tkinter as tk
from tkinter import ttk, messagebox
from versioning.version_manager import version_manager
from datetime import datetime

class VersionManagerGUI:
    def __init__(self, parent):
        self.parent = parent
        self.create_widgets()
        self.refresh_status()
        self.refresh_backup_list()

    def create_widgets(self):
        # 노트북(탭) 생성
        notebook = ttk.Notebook(self.parent)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 메인 탭
        main_frame = ttk.Frame(notebook)
        notebook.add(main_frame, text="버전 관리")
        self.create_main_tab(main_frame)
        
        # 백업 탭
        backup_frame = ttk.Frame(notebook)
        notebook.add(backup_frame, text="백업 관리")
        self.create_backup_tab(backup_frame)

    def create_main_tab(self, parent):
        title_label = ttk.Label(parent, text="🔖 버전 관리", font=("맑은 고딕", 16, "bold"))
        title_label.pack(pady=(0, 20))

        status_frame = ttk.LabelFrame(parent, text="현재 상태", padding=10)
        status_frame.pack(fill='x', pady=(0, 10))
        self.version_var = tk.StringVar()
        ttk.Label(status_frame, text="현재 버전:").grid(row=0, column=0, sticky='w', padx=(0, 10))
        ttk.Label(status_frame, textvariable=self.version_var, font=("맑은 고딕", 10, "bold")).grid(row=0, column=1, sticky='w')
        self.changes_var = tk.StringVar()
        ttk.Label(status_frame, text="변경사항:").grid(row=1, column=0, sticky='w', padx=(0, 10), pady=(5, 0))
        ttk.Label(status_frame, textvariable=self.changes_var).grid(row=1, column=1, sticky='w', pady=(5, 0))

        button_frame = ttk.LabelFrame(parent, text="버전 관리", padding=10)
        button_frame.pack(fill='x', pady=(0, 10))
        ttk.Button(button_frame, text="🏷️ 새 버전 태그", command=self.create_version_tag).grid(row=0, column=0, padx=5, pady=5)
        ttk.Button(button_frame, text="🔄 상태 새로고침", command=self.refresh_status).grid(row=0, column=1, padx=5, pady=5)

        history_frame = ttk.LabelFrame(parent, text="버전 히스토리", padding=10)
        history_frame.pack(fill='both', expand=True, pady=(0, 10))
        columns = ("버전", "날짜", "메시지", "커밋 해시")
        self.history_tree = ttk.Treeview(history_frame, columns=columns, show='headings', height=8)
        for col in columns:
            self.history_tree.heading(col, text=col)
            self.history_tree.column(col, width=150)
        scrollbar = ttk.Scrollbar(history_frame, orient='vertical', command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=scrollbar.set)
        self.history_tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        self.history_tree.bind('<Double-1>', self.show_diff_popup)
        self.restore_button = ttk.Button(history_frame, text="이 버전으로 복원", command=self.restore_selected_version)
        self.restore_button.pack(pady=5)

    def create_backup_tab(self, parent):
        title_label = ttk.Label(parent, text="💾 백업 관리", font=("맑은 고딕", 16, "bold"))
        title_label.pack(pady=(0, 20))

        # 백업 목록 프레임
        backup_list_frame = ttk.LabelFrame(parent, text="백업 목록", padding=10)
        backup_list_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        columns = ("날짜", "버전", "폴더명", "파일 수")
        self.backup_tree = ttk.Treeview(backup_list_frame, columns=columns, show='headings', height=10)
        for col in columns:
            self.backup_tree.heading(col, text=col)
            self.backup_tree.column(col, width=150)
        
        scrollbar = ttk.Scrollbar(backup_list_frame, orient='vertical', command=self.backup_tree.yview)
        self.backup_tree.configure(yscrollcommand=scrollbar.set)
        self.backup_tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # 백업 관련 버튼들
        backup_button_frame = ttk.Frame(parent)
        backup_button_frame.pack(fill='x', pady=(0, 10))
        ttk.Button(backup_button_frame, text="🔄 백업 목록 새로고침", command=self.refresh_backup_list).pack(side='left', padx=5)
        ttk.Button(backup_button_frame, text="📁 백업 폴더 열기", command=self.open_backup_folder).pack(side='left', padx=5)
        ttk.Button(backup_button_frame, text="📋 백업 정보 보기", command=self.show_backup_info).pack(side='left', padx=5)
        
        # 백업 정보 표시 영역
        self.backup_info_text = tk.Text(parent, height=8, wrap='word')
        backup_info_scrollbar = ttk.Scrollbar(parent, orient='vertical', command=self.backup_info_text.yview)
        self.backup_info_text.configure(yscrollcommand=backup_info_scrollbar.set)
        
        info_frame = ttk.LabelFrame(parent, text="백업 정보", padding=10)
        info_frame.pack(fill='both', expand=True)
        self.backup_info_text.pack(side='left', fill='both', expand=True, in_=info_frame)
        backup_info_scrollbar.pack(side='right', fill='y', in_=info_frame)

    def refresh_status(self):
        try:
            current_version = version_manager.get_current_version()
            self.version_var.set(current_version)
            git_status = version_manager.get_git_status()
            if git_status["has_changes"]:
                modified_count = len(git_status["modified"])
                staged_count = len(git_status["staged"])
                self.changes_var.set(f"수정됨: {modified_count}개, 스테이징됨: {staged_count}개")
            else:
                self.changes_var.set("변경사항 없음")
            for item in self.history_tree.get_children():
                self.history_tree.delete(item)
            history = version_manager.get_version_history()
            for item in history:
                date_str = item.get("date", "")
                self.history_tree.insert("", "end", values=(
                    item.get("version", ""),
                    date_str,
                    item.get("message", ""),
                    item.get("commit_hash", "")[:8] if item.get("commit_hash") else ""
                ))
        except Exception as e:
            messagebox.showerror("오류", f"상태를 새로고침하는 중 오류가 발생했습니다:\n{e}")

    def refresh_backup_list(self):
        """백업 목록을 새로고침합니다."""
        try:
            # 기존 항목 삭제
            for item in self.backup_tree.get_children():
                self.backup_tree.delete(item)
            
            # 백업 목록 가져오기
            backups = version_manager.get_backup_list()
            
            for backup in backups:
                # 파일 수 계산
                file_count = 0
                if os.path.exists(backup["path"]):
                    for root, dirs, files in os.walk(backup["path"]):
                        file_count += len([f for f in files if f.endswith('.py')])
                
                self.backup_tree.insert("", "end", values=(
                    backup["date"],
                    backup["version"],
                    backup["folder"],
                    file_count
                ))
                
        except Exception as e:
            messagebox.showerror("오류", f"백업 목록을 새로고침하는 중 오류가 발생했습니다:\n{e}")

    def open_backup_folder(self):
        """백업 폴더를 파일 탐색기에서 엽니다."""
        try:
            backup_dir = version_manager.backup_dir
            if os.path.exists(backup_dir):
                os.startfile(backup_dir)
            else:
                messagebox.showinfo("정보", "백업 폴더가 존재하지 않습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"백업 폴더를 여는 중 오류가 발생했습니다:\n{e}")

    def show_backup_info(self):
        """선택된 백업의 상세 정보를 표시합니다."""
        selected = self.backup_tree.selection()
        if not selected:
            messagebox.showwarning("경고", "백업 정보를 볼 항목을 선택하세요.")
            return
        
        item = self.backup_tree.item(selected[0])
        folder_name = item['values'][2]  # 폴더명
        
        try:
            backups = version_manager.get_backup_list()
            selected_backup = None
            for backup in backups:
                if backup["folder"] == folder_name:
                    selected_backup = backup
                    break
            
            if selected_backup:
                # 백업 정보 텍스트 영역에 표시
                self.backup_info_text.delete('1.0', tk.END)
                self.backup_info_text.insert('1.0', selected_backup["info"])
            else:
                messagebox.showwarning("경고", "선택된 백업 정보를 찾을 수 없습니다.")
                
        except Exception as e:
            messagebox.showerror("오류", f"백업 정보를 가져오는 중 오류가 발생했습니다:\n{e}")

    def create_version_tag(self):
        dialog = tk.Toplevel(self.parent)
        dialog.title("새 버전 태그 생성")
        dialog.geometry("400x250")
        dialog.transient(self.parent)
        dialog.grab_set()
        ttk.Label(dialog, text="버전 (예: 1.0.0):").pack(pady=(20, 5))
        version_entry = ttk.Entry(dialog, width=20)
        version_entry.pack(pady=(0, 10))
        version_entry.insert(0, "1.0.0")
        version_entry.focus()
        ttk.Label(dialog, text="메시지:").pack(pady=(10, 5))
        message_entry = ttk.Entry(dialog, width=50)
        message_entry.pack(pady=(0, 20))
        def create():
            version = version_entry.get().strip()
            message = message_entry.get().strip()
            if not version:
                messagebox.showwarning("경고", "버전을 입력해주세요.")
                return
            try:
                backup_info = version_manager.create_version_tag(version, message)
                messagebox.showinfo("성공", 
                    f"버전 태그가 생성되었습니다: v{version}\n\n"
                    f"백업 정보:\n"
                    f"- 백업 폴더: {backup_info['backup_folder']}\n"
                    f"- 백업된 파일 수: {backup_info['total_files']}개")
                dialog.destroy()
                self.refresh_status()
                self.refresh_backup_list()
            except Exception as e:
                messagebox.showerror("오류", f"버전 태그 생성 중 오류가 발생했습니다:\n{e}")
        ttk.Button(dialog, text="생성", command=create).pack(pady=10)
        dialog.bind('<Return>', lambda e: create())

    def show_diff_popup(self, event):
        selected = self.history_tree.selection()
        if not selected:
            return
        item = self.history_tree.item(selected[0])
        version = item['values'][0]
        # 이전 버전 찾기
        all_items = self.history_tree.get_children()
        idx = list(all_items).index(selected[0])
        if idx == len(all_items) - 1:
            messagebox.showinfo("정보", "이전 버전이 없습니다.")
            return
        prev_item = self.history_tree.item(all_items[idx + 1])
        prev_version = prev_item['values'][0]
        try:
            diff_text = version_manager.get_diff_between_versions(prev_version, version)
        except Exception as e:
            messagebox.showerror("오류", f"diff를 가져오는 중 오류 발생: {e}")
            return
        # 팝업으로 diff 보여주기
        diff_win = tk.Toplevel(self.parent)
        diff_win.title(f"{prev_version} ↔ {version} 변경점(diff)")
        diff_win.geometry("800x600")
        text = tk.Text(diff_win, wrap='none')
        text.insert('1.0', diff_text)
        text.config(state='disabled')
        text.pack(fill='both', expand=True)
        yscroll = ttk.Scrollbar(diff_win, orient='vertical', command=text.yview)
        text.configure(yscrollcommand=yscroll.set)
        yscroll.pack(side='right', fill='y')
        xscroll = ttk.Scrollbar(diff_win, orient='horizontal', command=text.xview)
        text.configure(xscrollcommand=xscroll.set)
        xscroll.pack(side='bottom', fill='x')

    def restore_selected_version(self):
        selected = self.history_tree.selection()
        if not selected:
            messagebox.showwarning("경고", "복원할 버전을 선택하세요.")
            return
        item = self.history_tree.item(selected[0])
        version = item['values'][0]
        if not version:
            messagebox.showwarning("경고", "유효한 버전이 아닙니다.")
            return
        try:
            version_manager.checkout_version(version)
            messagebox.showinfo("성공", f"{version} 버전으로 복원되었습니다.\n(Detached HEAD 상태입니다)")
            self.refresh_status()
        except Exception as e:
            messagebox.showerror("오류", f"복원 중 오류 발생: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("LISA - 버전 관리")
    VersionManagerGUI(root)
    root.mainloop() 