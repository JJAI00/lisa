import tkinter as tk
from tkinter import messagebox
import pandas as pd
from modules.renewal import renewal_search as rs
from core.data_loader import get_global_auth, safe_api_call
import time

class QuoteStatusModifier:
    def __init__(self):
        self.gc = get_global_auth()
        self.sheet_id = '1xjIjFe1Q9dq2zeQOkBSpbqIUOyXAxbyTO8lKN-r1lw0'
        self.worksheet_name = 'Sheet1'
        
        # 실제 quote_list 데이터가 저장된 시트 확인
        self._find_quote_sheet()
        
    def _find_quote_sheet(self):
        """실제 quote_list 데이터가 저장된 시트 찾기"""
        try:
            sh = self.gc.open_by_key(self.sheet_id)
            worksheets = sh.worksheets()
            
            print(f"시트 ID '{self.sheet_id}'에서 {len(worksheets)}개의 워크시트 발견:")
            for ws in worksheets:
                print(f"  - {ws.title}")
            
            # 모든 워크시트 확인
            for ws in worksheets:
                try:
                    # 각 워크시트에서 데이터 확인
                    all_values = ws.get_all_values()
                    if len(all_values) > 1:  # 헤더 + 데이터가 있는지 확인
                        headers = [h.strip() for h in all_values[0] if h.strip()]
                        print(f"  {ws.title} 헤더: {headers}")
                        
                        # Status 컬럼이 있는지 확인
                        has_status = any('status' in h.lower() or h == 'Status' for h in headers)
                        
                        # quote_list에 필요한 컬럼들이 있는지 확인
                        required_cols = ['일자', '영업사원', '대분류', '제품', '고객사명', '거래처명']
                        found_cols = [col for col in required_cols if col in headers]
                        
                        # Status 컬럼이 있는 quote_list를 우선적으로 찾기
                        if len(found_cols) >= 4 and has_status:
                            print(f"quote_list 데이터 발견 (Status 포함): {ws.title} (발견된 컬럼: {found_cols})")
                            self.worksheet_name = ws.title
                            return
                        elif len(found_cols) >= 4:
                            print(f"quote_list 데이터 발견 (Status 없음): {ws.title} (발견된 컬럼: {found_cols})")
                            # Status가 없는 경우도 저장하되, 나중에 Status 컬럼을 추가할 수 있도록 함
                            if not hasattr(self, 'backup_worksheet_name'):
                                self.backup_worksheet_name = ws.title
                except Exception as e:
                    print(f"워크시트 {ws.title} 확인 중 오류: {e}")
                    continue
            
            # Status가 포함된 시트를 찾지 못한 경우, 백업 시트 사용
            if hasattr(self, 'backup_worksheet_name'):
                print(f"Status 컬럼이 포함된 시트를 찾지 못해 백업 시트 사용: {self.backup_worksheet_name}")
                self.worksheet_name = self.backup_worksheet_name
            else:
                print(f"quote_list 데이터를 찾을 수 없어 기본 시트 사용: {self.worksheet_name}")
            
        except Exception as e:
            print(f"시트 찾기 실패: {e}")
        
    def update_status_in_sheet(self, row_data, new_status):
        """구글 시트에서 특정 행의 Status를 업데이트 (고유값으로 행 찾기)"""
        try:
            # 구글 시트에서 데이터 가져오기
            sh = self.gc.open_by_key(self.sheet_id)
            ws = sh.worksheet(self.worksheet_name)
            
            # 모든 데이터 가져오기
            all_values = ws.get_all_values()
            
            if len(all_values) > 1:
                # 헤더 행 확인
                headers = [h.strip() for h in all_values[0] if h.strip()]
                print(f"시트 '{self.worksheet_name}' 헤더: {headers}")
                
                # Status 컬럼 찾기
                status_col = None
                for i, header in enumerate(headers):
                    if 'status' in header.lower() or header == 'Status':
                        status_col = i + 1  # 구글 시트는 1부터 시작
                        print(f"Status 컬럼 발견: {header} (컬럼 {status_col})")
                        break
                
                # Status 컬럼을 찾지 못하면 에러
                if status_col is None:
                    print(f"Status 컬럼을 찾을 수 없습니다. 사용 가능한 컬럼: {headers}")
                    return False
                
                # 정확한 행 매칭을 위한 키 컬럼들 (모든 주요 필드 사용)
                key_cols = ['일자', '고객사명', '거래처명', '대분류', '제품', '수량', '판매 단가', '원가']
                
                # 키 컬럼이 모두 존재하는지 확인
                missing_cols = [col for col in key_cols if col not in headers]
                if missing_cols:
                    print(f"일부 키 컬럼이 시트에 없습니다: {missing_cols}")
                    # 누락된 컬럼 제거하고 진행
                    key_cols = [col for col in key_cols if col in headers]
                    print(f"사용 가능한 키 컬럼: {key_cols}")
                
                # 행 매칭 함수 (정확한 매칭)
                def row_matches(sheet_row):
                    try:
                        for col in key_cols:
                            sheet_value = str(sheet_row[headers.index(col)]).strip()
                            target_value = str(row_data.get(col, '')).strip()
                            
                            # 일자 컬럼의 경우 시간 정보 제거
                            if col == '일자':
                                if ' ' in sheet_value:
                                    sheet_value = sheet_value.split(' ')[0]
                                if ' ' in target_value:
                                    target_value = target_value.split(' ')[0]
                            
                            # 정확히 일치하는지 확인
                            if sheet_value != target_value:
                                return False
                        
                        return True
                    except (IndexError, ValueError) as e:
                        print(f"행 매칭 중 오류: {e}")
                        return False
                
                # 찾는 원본 데이터 출력
                print(f"찾는 원본 데이터:")
                for col in key_cols:
                    print(f"  {col}: {row_data.get(col, 'N/A')}")
                
                # 일치하는 행 찾기
                matched_rows = []
                for i, sheet_row in enumerate(all_values[1:], 2):  # 2번째 줄부터 데이터 (구글 시트는 1부터 시작)
                    if row_matches(sheet_row):
                        # 현재 값 확인
                        current_value = sheet_row[status_col - 1] if status_col - 1 < len(sheet_row) else "N/A"
                        print(f"정확히 일치하는 행 발견: 행 {i}")
                        print(f"  현재 Status: {current_value} -> 새 Status: {new_status}")
                        
                        # 시트 데이터 출력 (디버깅용)
                        sheet_data = {}
                        for col in key_cols:
                            col_idx = headers.index(col)
                            sheet_data[col] = sheet_row[col_idx] if col_idx < len(sheet_row) else 'N/A'
                        print(f"  시트 데이터: {sheet_data}")
                        
                        matched_rows.append(i)
                
                # 일치하는 행이 있으면 Status 업데이트
                if matched_rows:
                    if len(matched_rows) == 1:
                        # 정확히 하나의 행만 일치
                        row_idx = matched_rows[0]
                        ws.update_cell(row_idx, status_col, new_status)
                        print(f"Status 업데이트 완료: 시트 '{self.worksheet_name}', 행 {row_idx}, 컬럼 {status_col}, 새 값: {new_status}")
                        return True
                    else:
                        # 여러 행이 일치하는 경우 (예상치 못한 상황)
                        print(f"경고: {len(matched_rows)}개의 행이 일치합니다. 행 번호: {matched_rows}")
                        # 첫 번째 행만 업데이트
                        row_idx = matched_rows[0]
                        ws.update_cell(row_idx, status_col, new_status)
                        print(f"첫 번째 행의 Status 업데이트 완료: 행 {row_idx}, 새 값: {new_status}")
                        return True
                
                print(f"일치하는 행을 찾을 수 없습니다.")
                print(f"사용한 키 컬럼: {key_cols}")
                print(f"찾는 값: {row_data}")
                
                # 더 유연한 매칭 시도 (일자만으로 매칭)
                print("일자만으로 매칭 시도...")
                date_only_key = '일자'
                if date_only_key in headers:
                    date_col_idx = headers.index(date_only_key)
                    target_date = str(row_data.get(date_only_key, '')).strip()
                    
                    # 시간 정보 제거
                    if ' ' in target_date:
                        target_date = target_date.split(' ')[0]
                    
                    print(f"찾는 일자: {target_date}")
                    
                    for i, sheet_row in enumerate(all_values[1:], 2):
                        if len(sheet_row) > date_col_idx:
                            sheet_date = str(sheet_row[date_col_idx]).strip()
                            # 시간 정보 제거
                            if ' ' in sheet_date:
                                sheet_date = sheet_date.split(' ')[0]
                            
                            if sheet_date == target_date:
                                print(f"일자로 일치하는 행 발견: 행 {i}")
                                print(f"  시트 데이터: {sheet_row}")
                                print(f"  찾는 데이터: {row_data}")
                                
                                # Status 업데이트
                                ws.update_cell(i, status_col, new_status)
                                print(f"Status 업데이트 완료: 시트 '{self.worksheet_name}', 행 {i}, 컬럼 {status_col}, 새 값: {new_status}")
                                return True
                
                return False
            else:
                print(f"데이터가 없습니다: {len(all_values)}행")
                return False
                
        except Exception as e:
            print(f"Status 업데이트 실패: {e}")
            return False

def show_status_context_menu(event, tree, status_modifier):
    """우클릭 컨텍스트 메뉴 표시 (Status 메뉴 포함)"""
    try:
        # 선택된 항목 확인
        selected_item = tree.selection()
        if not selected_item:
            return
        
        # 선택된 항목의 데이터 가져오기
        item_data = tree.item(selected_item[0])
        values = item_data['values']
        columns = tree['columns']
        
        # 컨텍스트 메뉴 생성
        context_menu = tk.Menu(event.widget, tearoff=0)
        
        # Status 서브메뉴 생성
        status_menu = tk.Menu(context_menu, tearoff=0)
        status_values = ['20%', '40%', '60%', '80%', '100%', 'Drop']
        
        for status in status_values:
            status_menu.add_command(
                label=status,
                command=lambda s=status, item=selected_item[0]: update_status(tree, item, s, status_modifier)
            )
        
        # Status 메뉴 추가
        context_menu.add_cascade(label="Status", menu=status_menu)
        context_menu.add_separator()
        
        # 전자결재 보내기 메뉴 (항상 표시, Status 확인 후 처리)
        try:
            status_idx = list(columns).index('Status')
            status_value = values[status_idx]
            print(f"현재 Status 값: '{status_value}' (타입: {type(status_value)})")
            
            # Status 값 정규화 (숫자, 문자열, 퍼센트 등 다양한 형식 지원)
            normalized_status = str(status_value).strip()
            
            # 전자결재 보내기 메뉴 항상 추가
            context_menu.add_command(
                label="전자결재 보내기",
                command=lambda: send_to_approval_with_status_check(tree, selected_item[0], normalized_status)
            )
            print(f"전자결재 보내기 메뉴 추가됨 (Status: {normalized_status})")
        except Exception as e:
            print(f"Status 컬럼 인덱스 찾기 실패: {e}")
        
        # 메뉴 표시
        context_menu.tk_popup(event.x_root, event.y_root)
        
    except Exception as e:
        messagebox.showerror("오류", f"컨텍스트 메뉴 생성 실패: {e}")

def update_status(tree, selected_item, new_status, status_modifier):
    """선택된 항목의 Status를 업데이트 (같은 ID를 가진 모든 행 포함)"""
    try:
        # 선택된 항목의 데이터 가져오기
        item_data = tree.item(selected_item)
        values = item_data['values']
        columns = tree['columns']
        
        # ID 컬럼 인덱스 찾기
        try:
            id_idx = list(columns).index('ID')
            selected_id = values[id_idx] if id_idx < len(values) else None
        except ValueError:
            # ID 컬럼이 없는 경우 기존 방식 사용
            selected_id = None
        
        if selected_id:
            print(f"선택된 ID: {selected_id}")
            
            # 같은 ID를 가진 모든 행 찾기
            all_items = tree.get_children()
            same_id_items = []
            
            for item in all_items:
                item_values = tree.item(item)['values']
                if id_idx < len(item_values) and item_values[id_idx] == selected_id:
                    same_id_items.append(item)
            
            print(f"같은 ID를 가진 행 개수: {len(same_id_items)}")
            
            # 모든 같은 ID 행의 Status 업데이트
            updated_count = 0
            for item in same_id_items:
                item_values = tree.item(item)['values']
                
                # 행 데이터를 딕셔너리로 만들기
                row_data = {}
                for i, col in enumerate(columns):
                    if i < len(item_values):
                        row_data[col] = item_values[i]
                
                print(f"업데이트할 행 데이터: {row_data}")
                
                # Status 업데이트
                success = status_modifier.update_status_in_sheet(row_data, new_status)
                
                if success:
                    # 트리뷰에서도 값 업데이트
                    try:
                        status_idx = list(columns).index('Status')
                        item_values[status_idx] = new_status
                        tree.item(item, values=item_values)
                        updated_count += 1
                    except Exception as e:
                        print(f"Status 컬럼 인덱스 찾기 실패(트리뷰): {e}")
                else:
                    print(f"행 업데이트 실패: {row_data}")
            
            if updated_count > 0:
                messagebox.showinfo("성공", f"같은 ID를 가진 {updated_count}개 행의 Status가 {new_status}로 업데이트되었습니다.")
            else:
                messagebox.showerror("실패", "Status 업데이트에 실패했습니다.")
        else:
            # ID가 없는 경우 기존 방식 사용 (단일 행만 업데이트)
            print("ID 컬럼이 없어 단일 행만 업데이트합니다.")
            
            # 선택된 행의 데이터를 딕셔너리로 만들기
            row_data = {}
            for i, col in enumerate(columns):
                if i < len(values):
                    row_data[col] = values[i]
            
            print(f"선택된 행 데이터: {row_data}")
            
            # Status 업데이트 (고유값으로 행 찾기)
            success = status_modifier.update_status_in_sheet(row_data, new_status)
            
            if success:
                # 트리뷰에서도 값 업데이트 (Status 컬럼 인덱스를 동적으로 찾음)
                try:
                    status_idx = list(columns).index('Status')
                    values[status_idx] = new_status
                    tree.item(selected_item, values=values)
                except Exception as e:
                    print(f"Status 컬럼 인덱스 찾기 실패(트리뷰): {e}")
                
                messagebox.showinfo("성공", f"Status가 {new_status}로 업데이트되었습니다.")
            else:
                messagebox.showerror("실패", "Status 업데이트에 실패했습니다.")
            
    except Exception as e:
        messagebox.showerror("오류", f"Status 업데이트 실패: {e}")

def send_to_approval_with_status_check(tree, selected_item, current_status):
    """Status를 확인하고 전자결재 보내기 처리"""
    try:
        # Status가 80%인지 확인
        if current_status in ["80", "80%", "80.0", "80.0%"]:
            # 80%인 경우 바로 전자결재 실행
            from lisa_daou_sign import send_to_approval as original_send_to_approval
            original_send_to_approval(tree, selected_item)
        else:
            # 80%가 아닌 경우 안내 메시지 표시
            messagebox.showwarning(
                "Status 확인", 
                f"현재 Status가 '{current_status}'입니다.\n\n"
                "전자결재를 진행하려면 Status를 80%로 수정 후 다시 시도해 주세요.\n\n"
                "Status 메뉴에서 80%를 선택하여 변경할 수 있습니다."
            )
    except Exception as e:
        messagebox.showerror("오류", f"전자결재 전송 실패: {e}")

def send_to_approval(tree, selected_item):
    """전자결재 보내기 함수 (기존 기능)"""
    try:
        from lisa_daou_sign import send_to_approval as original_send_to_approval
        original_send_to_approval(tree, selected_item)
    except Exception as e:
        messagebox.showerror("오류", f"전자결재 전송 실패: {e}")

if __name__ == "__main__":
    # 테스트용 코드
    status_modifier = QuoteStatusModifier()
    print("견적서 Status 수정 모듈이 준비되었습니다.") 