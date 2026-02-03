# license_certificate_info.py

import tkinter as tk
from .license_certificate_BorisFX import open_borisfx_dialog
from .license_certificate_Marmoset import open_marmoset_dialog
from .license_certificate_Maxon import open_maxon_dialog
from .license_certificate_VideoCopilot import open_videocopilot_dialog
from .license_certificate_Foundry import open_foundry_dialog

def open_dialog(parent_frame, username=None):
    username = username or '테스트유저'
    # 상단 헤더
    header_frame = tk.Frame(parent_frame, bg="#F7F9FB", height=60)
    header_frame.pack(fill='x', pady=(0, 20))
    header_frame.pack_propagate(False)
    
    # 헤더 제목
    title_label = tk.Label(header_frame, text="🔐 인증서 관리", 
                          font=("맑은 고딕", 16, "bold"), 
                          fg="#1F2937", bg="#F7F9FB")
    title_label.pack(side='left', padx=20, pady=15)
    
    # 사용자 정보 (우측)
    user_label = tk.Label(header_frame, text=f"👤 {username}", 
                         font=("맑은 고딕", 11), 
                         fg="#6B7280", bg="#F7F9FB")
    user_label.pack(side='right', padx=20, pady=15)

    # 인증서 카드 섹션
    card_section = tk.Frame(parent_frame, bg="#F7F9FB")
    card_section.pack(fill='both', expand=True, padx=20, pady=(0, 20))
    
    # 카드 헤더
    card_header = tk.Frame(card_section, bg="white", relief="solid", bd=1)
    card_header.pack(fill='x', pady=(0, 1))
    tk.Label(card_header, text="📜 인증서 만들기", font=("맑은 고딕", 11, "bold"), 
            fg="#1F2937", bg="white").pack(side='left', padx=15, pady=10)
    
    # 카드 내용
    card_content = tk.Frame(card_section, bg="white", relief="solid", bd=1)
    card_content.pack(fill='both', expand=True, padx=20, pady=20)
    
    # 버튼 그리드 (2열로 배치)
    button_frame = tk.Frame(card_content, bg="white")
    button_frame.pack(expand=True)
    
    # 버튼 정의 (아이콘, 텍스트, 색상, 명령어)
    buttons = [
        ("🎨", "BorisFX 인증서 만들기", "#8B5CF6", lambda: open_borisfx_dialog(parent_frame)),
        ("🔧", "Foundry 인증서 만들기", "#3B82F6", lambda: open_foundry_dialog(parent_frame)),
        ("🎮", "Marmoset 인증서 만들기", "#10B981", lambda: open_marmoset_dialog(parent_frame)),
        ("🎬", "Maxon 인증서 만들기", "#F59E0B", lambda: open_maxon_dialog(parent_frame)),
        ("🎥", "VideoCopilot 인증서 만들기", "#EF4444", lambda: open_videocopilot_dialog(parent_frame))
    ]
    
    # 2열 그리드로 버튼 배치
    for i, (icon, text, color, command) in enumerate(buttons):
        row = i // 2
        col = i % 2
        
        # 버튼 프레임
        btn_container = tk.Frame(button_frame, bg="white")
        btn_container.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")
        
        # 버튼
        btn = tk.Button(btn_container, 
                       text=f"{icon}\n{text}", 
                       command=command,
                       font=("맑은 고딕", 10, "bold"),
                       bg=color,
                       fg="white",
                       relief="flat",
                       width=20,
                       height=4,
                       cursor="hand2")
        btn.pack(expand=True, fill="both", padx=5, pady=5)
        
        # hover 효과
        def create_hover_effect(button, base_color):
            def on_enter(e):
                # 색상을 약간 어둡게
                if base_color == "#8B5CF6":
                    button.configure(bg="#7C3AED")
                elif base_color == "#3B82F6":
                    button.configure(bg="#2563EB")
                elif base_color == "#10B981":
                    button.configure(bg="#059669")
                elif base_color == "#F59E0B":
                    button.configure(bg="#D97706")
                elif base_color == "#EF4444":
                    button.configure(bg="#DC2626")
            
            def on_leave(e):
                button.configure(bg=base_color)
            
            button.bind('<Enter>', on_enter)
            button.bind('<Leave>', on_leave)
        
        create_hover_effect(btn, color)
    
    # 그리드 가중치 설정
    button_frame.grid_columnconfigure(0, weight=1)
    button_frame.grid_columnconfigure(1, weight=1)
    button_frame.grid_rowconfigure(0, weight=1)
    button_frame.grid_rowconfigure(1, weight=1)
    button_frame.grid_rowconfigure(2, weight=1)


if __name__ == '__main__':
    # 메인 윈도우
    root = tk.Tk()
    root.title("License Certificate Info")

    # 메인 버튼
    btn = tk.Button(root, text="인증서 만들기", command=lambda: open_dialog(root), width=25, height=2)
    btn.pack(padx=20, pady=20)

    root.mainloop()
