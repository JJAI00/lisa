import tkinter as tk
import pandas as pd
from core import data_loader


def get_salesperson_from_renewal(customer_name):
    """renewal_list에서 해당 고객사의 영업사원 정보 추출"""
    try:
        # data_loader에서 renewal 데이터 가져오기
        loader = data_loader.DataLoader('google_sheet_key/renewal-bot-463605-b8f41de8fbdb.json', 
                                      ['https://www.googleapis.com/auth/spreadsheets.readonly'])
        renewal_data = loader.get_renewal_data()
        
        if renewal_data is not None and not renewal_data.empty:
            # 해당 고객사의 renewal 데이터 찾기
            customer_renewals = renewal_data[renewal_data['고객사명'] == customer_name]
            if not customer_renewals.empty:
                # 가장 최근 데이터의 영업사원 반환 (일자 기준으로 정렬)
                if '일자' in customer_renewals.columns:
                    customer_renewals = customer_renewals.sort_values('일자', ascending=False)
                return customer_renewals.iloc[0].get('영업사원', '')
        return ''
    except Exception as e:
        print(f"영업사원 정보 조회 오류: {e}")
        return ''

def show_customer_detail(parent, row):
    """
    별도 창을 띄워 고객 상세 정보를 표시합니다.
    parent: Tk or Toplevel
    row: pandas Series 또는 dict
    """
    detail = tk.Toplevel(parent)
    detail.title(f"상세 정보: {row['고객사명']}")
    detail.geometry("600x900")
    detail.configure(bg="#F7F9FB")
    detail.resizable(False, False)

    # 영업사원 정보 가져오기
    salesperson = get_salesperson_from_renewal(row['고객사명'])

    # 금액 관련 계산
    price    = row.get('2024년 판매 단가', 0)
    sale_sum = row.get('2024년 판매 합계', 0)
    cost     = row.get('2024년 판매 원가', 0)
    cost_sum = row.get('2024년 판매원가계', 0)
    profit   = price - cost
    margin   = f"{profit/price:.2%}" if price else ''
    expiry = ''
    if '만료일' in row and isinstance(row['만료일'], pd.Timestamp) and not pd.isna(row['만료일']):
        expiry = row['만료일'].strftime('%Y-%m-%d')
    elif '만료일' in row and isinstance(row['만료일'], str) and row['만료일'].strip():
        expiry = row['만료일']

    fields = [
        ("고객사명",    row['고객사명']),
        ("거래처명",    row.get('거래처명','')),
        ("영업사원",    salesperson),
        ("담당자명",    row.get('담당자명','')),
        ("이메일",     row.get('이메일','')),
        ("연락처",     row.get('연락처','')),
        ("핸드폰",     row.get('핸드폰','')),
        ("주소",       row.get('주소','')),
        ("제품",       row['제품']),
        ("수량",       row['수량']),
        ("만료일",     expiry),
        ("24년 판매 단가",   f"{row.get('판매 단가', 0):,}원"),
        ("24년 판매 합계",   f"{row.get('판매 합계', 0):,}원"),
        ("24년 판매 원가",   f"{row.get('원가', 0):,}원"),
        ("24년 원가계",      f"{row.get('원가계', 0):,}원"),
        ("이익액",           f"{row.get('이익액', 0):,}원"),
        ("이익율",      row['이익율'])
    ]

    create_detail_widgets(detail, fields)

    detail.transient(parent)
    detail.grab_set()
    detail.wait_window()


def create_detail_widgets(detail, fields):
    """상세 정보 위젯 생성"""
    # 상단 헤더
    header_frame = tk.Frame(detail, bg="#F7F9FB", height=60)
    header_frame.pack(fill='x', pady=(0, 20))
    header_frame.pack_propagate(False)
    
    # 헤더 제목
    title_label = tk.Label(header_frame, text="👤 고객사 상세 정보", 
                          font=("맑은 고딕", 16, "bold"), 
                          fg="#1F2937", bg="#F7F9FB")
    title_label.pack(side='left', padx=20, pady=15)

    # 메인 카드
    main_card = tk.Frame(detail, bg="white", relief="solid", bd=1)
    main_card.pack(fill='both', expand=True, padx=20, pady=(0, 20))
    
    # 카드 헤더
    card_header = tk.Frame(main_card, bg="#F8FAFC", height=40)
    card_header.pack(fill='x')
    card_header.pack_propagate(False)
    tk.Label(card_header, text="📋 상세 정보", font=("맑은 고딕", 11, "bold"), 
            fg="#374151", bg="#F8FAFC").pack(side='left', padx=15, pady=10)
    
    # 카드 내용
    card_content = tk.Frame(main_card, bg="white")
    card_content.pack(fill='both', expand=True, padx=20, pady=20)
    
    # 필드들을 섹션별로 그룹화
    sections = [
        ("고객사 정보", fields[:8]),
        ("제품 정보", fields[8:11]),
        ("금액 정보", fields[11:])
    ]
    
    for section_title, section_fields in sections:
        # 섹션 제목
        section_label = tk.Label(card_content, text=f"📌 {section_title}", 
                                font=("맑은 고딕", 12, "bold"), 
                                fg="#1F2937", bg="white", anchor='w')
        section_label.pack(fill='x', pady=(15, 10))
        
        # 섹션 필드들
        for label, value in section_fields:
            field_frame = tk.Frame(card_content, bg="white")
            field_frame.pack(fill='x', pady=3)
            
            # 라벨
            label_widget = tk.Label(field_frame, text=f"{label}:", 
                                   font=("맑은 고딕", 10, "bold"), 
                                   fg="#374151", bg="white", width=15, anchor='w')
            label_widget.pack(side='left')
            
            # 값
            value_widget = tk.Label(field_frame, text=value if value else "(없음)", 
                                   font=("맑은 고딕", 10), 
                                   fg="#6B7280" if not value else "#1F2937", 
                                   bg="white", anchor='w')
            value_widget.pack(side='left', padx=(10, 0), fill='x', expand=True)
            
            # 구분선 (마지막 필드가 아닌 경우)
            if label != section_fields[-1][0]:
                separator = tk.Frame(field_frame, height=1, bg="#E5E7EB")
                separator.pack(fill='x', pady=(5, 0))

    # 버튼 프레임
    button_frame = tk.Frame(detail, bg="#F7F9FB")
    button_frame.pack(fill='x', padx=20, pady=(0, 20))
    
    # 닫기 버튼
    close_btn = tk.Button(button_frame, text="❌ 닫기", command=detail.destroy,
                          font=("맑은 고딕", 11), bg="#6B7280", fg="white",
                          relief="flat", padx=20, pady=8, cursor="hand2")
    close_btn.pack(side='right')
    close_btn.bind('<Enter>', lambda e: close_btn.configure(bg="#374151"))
    close_btn.bind('<Leave>', lambda e: close_btn.configure(bg="#6B7280"))
