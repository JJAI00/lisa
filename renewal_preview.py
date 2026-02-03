import os
import calendar
import tkinter as tk
from tkinter import messagebox, ttk
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

import renewal_search as rs
import renewal_email as res_email
import renewal_quote as res_quote
import renewal_quote_multi as res_quote_multi
from renewal_email_multi import send_multi_emails
from sendlog import show_logs, KEY_FILE, SCOPES, LOG_SHEET_ID, LOG_SHEET_NAME
from customer_details import show_customer_detail

class RenewalMailerGUI:
    """메인 GUI (미리보기 + 액션)"""
    def __init__(self, root, username=None):
        self.root     = root
        self.username = username or '테스트유저'
        if hasattr(root, 'title'):
            root.title(f"LISA v1.5 - {self.username}")
        if hasattr(root, 'geometry'):
            root.geometry("1200x650")

        try:
            today = datetime.today().date()
            self.raw_df = rs.get_filtered_data_range(
                today - timedelta(days=3650),
                today + timedelta(days=3650)
            )
        except Exception as e:
            messagebox.showwarning("경고", f"데이터 로드 실패: {e}")
            self.raw_df = pd.DataFrame()

        self.current_df   = None
        self.last_start   = None
        self.last_end     = None
        self.sort_column  = None
        self.sort_reverse = False
        self.log_df       = pd.DataFrame(columns=['고객사명','제품','만료일','액션','일시'])
        self.rows         = []

        self.setup_widgets()
        self.load_filter_options()
        self._load_logs()
        self.load_this_month_preview()

    def _load_logs(self):
        """로그 시트를 불러와 self.log_df에 저장"""
        try:
            creds = Credentials.from_service_account_file(KEY_FILE, scopes=SCOPES)
            gc = gspread.authorize(creds)
            ws = gc.open_by_key(LOG_SHEET_ID).worksheet(LOG_SHEET_NAME)
            logs = ws.get_all_records()
            self.log_df = pd.DataFrame(logs)
            self.log_df.columns = self.log_df.columns.str.strip()
            if '일시' in self.log_df.columns:
                self.log_df['일시'] = pd.to_datetime(
                    self.log_df['일시'], format='%Y-%m-%d %H:%M', errors='coerce'
                )
                self.log_df.sort_values('일시', ascending=False, inplace=True)
        except Exception:
            self.log_df = pd.DataFrame(columns=['고객사명','제품','만료일','액션','일시'])

    def setup_widgets(self):
        # 상단: 날짜 선택 + 기간 버튼
        top = tk.Frame(self.root)
        top.pack(fill='x', pady=5)

        tk.Label(top, text="조회 기준일:").pack(side='left')
        # 복사본과 동일한 기본 DateEntry 사용 :contentReference[oaicite:2]{index=2}
        self.date_entry = DateEntry(top, width=12, date_pattern='yyyy-mm-dd')
        self.date_entry.set_date(datetime.today())
        self.date_entry.pack(side='left', padx=5)
        self.date_entry.bind("<<DateEntrySelected>>", lambda e: self.load_this_month_preview())

        for text, cmd in [
            ("이번달",   self.load_this_month_preview),
            ("다음달",   self.load_next_month_preview),
            ("미래 1년", self.load_one_year_preview),
            ("과거 1년", self.load_past_year_preview),
        ]:
            tk.Button(top, text=text, command=cmd).pack(side='left', padx=5)

        tk.Button(top, text="로그 보기", command=lambda: show_logs(self.root)).pack(side='right', padx=5)

        # 중단: 필터 및 액션 버튼
        mid = tk.Frame(self.root)
        mid.pack(fill='x', pady=5)

        tk.Label(mid, text="대분류:").pack(side='left')
        self.category_cb = ttk.Combobox(mid, state='readonly', width=12)
        self.category_cb.pack(side='left', padx=5)
        self.category_cb.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        tk.Label(mid, text="영업사원:").pack(side='left', padx=(20,0))
        self.sales_cb = ttk.Combobox(mid, state='readonly', width=12)
        self.sales_cb.pack(side='left', padx=5)
        self.sales_cb.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        tk.Label(mid, text="고객사명:").pack(side='left', padx=(20,0))
        self.search_entry = tk.Entry(mid, width=15)
        self.search_entry.pack(side='left', padx=5)

        tk.Label(mid, text="거래처명:").pack(side='left', padx=(20,0))
        self.client_search_entry = tk.Entry(mid, width=15)
        self.client_search_entry.pack(side='left', padx=5)

        tk.Button(mid, text="검색", command=self.apply_filters).pack(side='left', padx=5)
        tk.Button(mid, text="리로드", command=self.reload_current).pack(side='left', padx=5)
        tk.Button(mid, text="메일 작성", command=self.send_selected_email).pack(side='left', padx=20)
        tk.Button(mid, text="견적서 첨부", command=self.quote_selected).pack(side='left', padx=5)

        # 하단: 결과 테이블
        self.cols = (
            "액션","계산서","고객사명","거래처명","영업담당자",
            "대분류","제품","수량","만료일",
            "판매 단가","판매 합계","원가","원가계","이익액","이익율"
        )
        container = tk.Frame(self.root)
        container.pack(fill='both', expand=True)
        vsb = ttk.Scrollbar(container, orient='vertical'); vsb.pack(side='right', fill='y')
        hsb = ttk.Scrollbar(container, orient='horizontal'); hsb.pack(side='bottom', fill='x')

        self.tree = ttk.Treeview(
            container, columns=self.cols, show='headings', selectmode='browse',
            yscrollcommand=vsb.set, xscrollcommand=hsb.set
        )
        vsb.config(command=self.tree.yview); hsb.config(command=self.tree.xview)

        for c in self.cols:
            self.tree.heading(c, text=c, command=lambda col=c: self.sort_by_column(col))
            if c in ("고객사명","영업담당자","대분류","제품"):
                anchor = 'w'
            elif c in ("판매 단가","판매 합계","원가","원가계","이익액"):
                anchor = 'e'
            else:
                anchor = 'center'
            self.tree.column(c, width=120, anchor=anchor, stretch=False)

        self.tree.pack(fill='both', expand=True)
        self.tree.bind("<Double-1>", self.show_detail)

    def load_filter_options(self):
        try:
            today = datetime.today().date()
            df_all = rs.get_filtered_data_range(
                today - timedelta(days=3650),
                today + timedelta(days=3650)
            )
            cats = sorted(df_all['대분류'].dropna().unique().tolist())
            self.category_cb['values'] = ['전체'] + cats
            self.category_cb.current(0)
            sales = sorted(df_all['영업 담당자'].dropna().unique().tolist())
        except:
            self.category_cb['values'] = ['전체']; self.category_cb.current(0); sales = []
        self.sales_cb['values'] = ['전체'] + sales
        self.sales_cb.current(0)

    def _filter_cached_data(self, start, end):
        self.last_start, self.last_end = start, end
        df = self.raw_df
        mask = (df['만료일'] >= start) & (df['만료일'] <= end)
        sel_df = df.loc[mask]
        self.current_df = sel_df.drop_duplicates(subset=['고객사명','제품','만료일']) if not sel_df.empty else sel_df
        self.populate_tree(self.current_df)

    def load_this_month_preview(self):
        self.clear_search_fields()
        base = self.date_entry.get_date()
        start = base.replace(day=1)
        last_day = calendar.monthrange(base.year, base.month)[1]
        end = base.replace(day=last_day)
        self._filter_cached_data(start, end)

    def load_next_month_preview(self):
        self.clear_search_fields()
        base = self.date_entry.get_date()
        year = base.year + (base.month // 12)
        month = (base.month % 12) + 1
        start = base.replace(year=year, month=month, day=1)
        last_day = calendar.monthrange(year, month)[1]
        end = start.replace(day=last_day)
        self._filter_cached_data(start, end)

    def load_one_year_preview(self):
        self.clear_search_fields()
        s, e = rs.period_one_year(self.date_entry.get_date())
        self._filter_cached_data(s, e)

    def load_past_year_preview(self):
        self.clear_search_fields()
        s, e = rs.period_past_year(self.date_entry.get_date())
        self._filter_cached_data(s, e)

    def clear_search_fields(self):
        self.category_cb.current(0)
        self.sales_cb.current(0)
        self.search_entry.delete(0, tk.END)
        self.client_search_entry.delete(0, tk.END)

    def reload_current(self):
        if self.last_start is None:
            return
        self._load_data(self.last_start, self.last_end)

    def apply_filters(self):
        if self.current_df is None:
            return
        df = self.current_df.copy()
        if self.category_cb.get() != '전체':
            df = df[df['대분류'] == self.category_cb.get()]
        if self.sales_cb.get() != '전체':
            df = df[df['영업 담당자'] == self.sales_cb.get()]
        kw = self.search_entry.get().strip()
        ckw = self.client_search_entry.get().strip()
        if kw:
            df = df[df['고객사명'].str.contains(kw, na=False)]
        if ckw:
            df = df[df['거래처명'].str.contains(ckw, na=False)]
        self.populate_tree(df)

    def populate_tree(self, df):
        self.tree.delete(*self.tree.get_children())
        self.rows = []
        for idx, row in df.iterrows():
            expiry = row['만료일'].strftime('%Y-%m-%d')
            if not self.log_df.empty:
                mask = (
                    (self.log_df['고객사명'] == row['고객사명']) &
                    (self.log_df['제품']    == row['제품']) &
                    (self.log_df['만료일']  == expiry)
                )
                acts = self.log_df.loc[mask, '액션']
            else:
                acts = pd.Series(dtype=object)
            mark = acts.iloc[-1] if len(acts) > 0 else ''

            uc = row.get('실 매입 단가')  if pd.notna(row.get('실 매입 단가'))  else row.get('원가',0)
            sc = row.get('실 매입 합계') if pd.notna(row.get('실 매입 합계')) else row.get('원가계',0)
            pv = row.get('실 이익액')    if pd.notna(row.get('실 이익액'))    else row.get('이익액',0)
            pr = row.get('실 이익율')    if pd.notna(row.get('실 이익율'))    else row.get('이익율',0)

            unit_cost_str = f"{uc:,}" if isinstance(uc, (int, float)) else str(uc)
            sum_cost_str  = f"{sc:,}" if isinstance(sc, (int, float)) else str(sc)
            profit_str    = f"{pv:,}" if isinstance(pv, (int, float)) else str(pv)
            rate_str      = pr if isinstance(pr, str) else f"{pr:.2%}"

            vals = (
                mark, row.get('계산서 발행일',''),
                row['고객사명'], row.get('거래처명',''), row.get('영업 담당자',''),
                row.get('대분류',''), row['제품'], row['수량'], expiry,
                f"{row.get('판매 단가',0):,}", f"{row.get('판매 합계',0):,}",
                unit_cost_str, sum_cost_str, profit_str, rate_str
            )
            # iid는 원본 인덱스 문자열 유지
            self.tree.insert('', 'end', iid=str(idx), values=vals)
            self.rows.append(row)

    def sort_by_column(self, col):
        if col == '액션': return
        df = self.current_df.copy()
        rev = (col == self.sort_column) and not self.sort_reverse
        self.sort_column, self.sort_reverse = col, rev
        nums = ("수량","판매 단가","판매 합계","원가","원가계")
        if col in nums or col in ("만료일","계산서 발행일","이익액","이익율"):
            df = df.sort_values(col, ascending=not rev)
        else:
            df = df.sort_values(col, ascending=not rev, key=lambda x: x.fillna('').astype(str))
        self.populate_tree(df)

    def show_detail(self, event):
        sel = self.tree.selection()
        if not sel: return
        pos = self.tree.index(sel[0])
        if 0 <= pos < len(self.rows):
            show_customer_detail(self.root, self.rows[pos])

    def send_selected_email(self):
        sel = self.tree.selection()
        if len(sel) != 1:
            messagebox.showwarning("경고", "메일 작성은 한 행만 선택해야 합니다.")
            return
        iid = sel[0]
        pos = self.tree.index(iid)
        row = self.rows[pos]
        sel_month = row['만료일'].month
        client = row.get('거래처명','')
        same = [j for j,r in enumerate(self.rows)
                if r['만료일'].month==sel_month and r.get('거래처명','')==client]
        if len(same) >= 2:
            dlg = tk.Toplevel(self.root); dlg.title("유통건 확인")
            tk.Label(dlg, text="2건 이상이 같은달에 있는 건입니다. 어떻게 하시겠어요?", padx=20, pady=10).pack()
            def single():
                dlg.destroy()
                res_email.send_emails([row], self.username)
                self.tree.set(iid, '액션', '메일 작성')
                messagebox.showinfo("완료","메일 작성 로그가 기록되었습니다.")
            def multi():
                dlg.destroy()
                lst = [self.rows[j] for j in same]
                send_multi_emails(lst, self.username)
                for j in same:
                    child_iid = self.tree.get_children()[j]
                    self.tree.set(child_iid, '액션', '메일 작성')
                messagebox.showinfo("완료", f"{len(same)}건 메일 작성 로그가 기록되었습니다.")
            frm = tk.Frame(dlg); frm.pack(pady=10)
            tk.Button(frm, text="1건만 발송", width=12, command=single).pack(side='left', padx=5)
            tk.Button(frm, text="2건 이상 발송", width=12, command=multi).pack(side='right', padx=5)
            dlg.transient(self.root); dlg.grab_set(); self.root.wait_window(dlg)
            return
        res_email.send_emails([row], self.username)
        self.tree.set(iid, '액션', '메일 작성')
        messagebox.showinfo("완료","메일 작성 로그가 기록되었습니다.")

    def quote_selected(self):
        sel = self.tree.selection()
        if len(sel) != 1:
            messagebox.showwarning("경고", "견적서 첨부는 한 행만 선택해야 합니다.")
            return
        iid = sel[0]
        pos = self.tree.index(iid)
        row = self.rows[pos]
        # 유통유무 체크
        if row.get('유통유무','') == '유통':
            messagebox.showwarning("경고", "유통 건은 직접 견적서를 작성해 주세요")
            return
        # Foundry 체크
        if row.get('대분류','') == 'Foundry':
            messagebox.showwarning("경고", "Foundry 제품은 영업사원이 직접 견적서를 작성해야 합니다")
            return
        sel_month = row['만료일'].month
        client = row.get('고객사명','')
        same = [j for j,r in enumerate(self.rows)
                if r['만료일'].month==sel_month and r.get('고객사명','')==client]
        if len(same) >= 2:
            dlg = tk.Toplevel(self.root); dlg.title("견적서 작성 방식 선택")
            tk.Label(dlg, text="2건 이상이 같은달에 있는 고객사 입니다. 어떻게 하시겠어요?", padx=20, pady=10).pack()
            def single_q():
                dlg.destroy()
                res_quote.create_and_send_quote(row, self.username)
                self.tree.set(iid, '액션', '메일+견적서')
                messagebox.showinfo("완료","견적서 작성 로그가 기록되었습니다.")
            def multi_q():
                dlg.destroy()
                lst = [self.rows[j] for j in same]
                res_quote_multi.create_and_send_quotes(lst, self.username)
                for j in same:
                    child_iid = self.tree.get_children()[j]
                    self.tree.set(child_iid, '액션', '메일+견적서')
                messagebox.showinfo("완료", f"{len(same)}건 견적서 작성 로그가 기록되었습니다.")
            frm = tk.Frame(dlg); frm.pack(pady=10)
            tk.Button(frm, text="1건만 작성", width=12, command=single_q).pack(side='left', padx=5)
            tk.Button(frm, text="2건 이상 작성", width=12, command=multi_q).pack(side='right', padx=5)
            dlg.transient(self.root); dlg.grab_set(); self.root.wait_window(dlg)
            return
        res_quote.create_and_send_quote(row, self.username)
        self.tree.set(iid, '액션', '메일+견적서')
        messagebox.showinfo("완료","견적서 작성 로그가 기록되었습니다.")

if __name__ == '__main__':
    root = tk.Tk()
    RenewalMailerGUI(root, username="테스트유저")
    root.mainloop()
