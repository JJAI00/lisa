"""
LISA 프로젝트 전체 분석 및 Excel 기획서 생성 스크립트
모든 Python 파일을 분석하여 체계적인 기획서 형태의 Excel 문서를 생성합니다.
"""

import os
import ast
import pandas as pd
from pathlib import Path
from datetime import datetime
import json

class PythonFileAnalyzer:
    """Python 파일 분석기"""
    
    def __init__(self, project_root="."):
        self.project_root = Path(project_root)
        self.analysis_results = []
        
    def analyze_file(self, file_path):
        """개별 Python 파일 분석"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
                tree = ast.parse(content, filename=str(file_path))
            
            file_info = {
                'file_path': str(file_path.relative_to(self.project_root)),
                'file_name': file_path.name,
                'module_name': file_path.stem,
                'lines_of_code': len(content.splitlines()),
                'docstring': ast.get_docstring(tree) or '',
                'classes': [],
                'functions': [],
                'imports': [],
                'constants': [],
                'category': self._categorize_file(file_path),
                'purpose': self._infer_purpose(file_path, tree),
                'dependencies': []
            }
            
            # AST 분석
            for node in ast.walk(tree):
                # 클래스 분석
                if isinstance(node, ast.ClassDef):
                    class_info = {
                        'name': node.name,
                        'docstring': ast.get_docstring(node) or '',
                        'methods': [m.name for m in node.body if isinstance(m, ast.FunctionDef)],
                        'bases': [self._get_name(base) for base in node.bases],
                        'line_no': node.lineno
                    }
                    file_info['classes'].append(class_info)
                
                # 함수 분석 (클래스 외부)
                elif isinstance(node, ast.FunctionDef) and not any(isinstance(parent, ast.ClassDef) for parent in ast.walk(tree)):
                    func_info = {
                        'name': node.name,
                        'docstring': ast.get_docstring(node) or '',
                        'args': [arg.arg for arg in node.args.args],
                        'line_no': node.lineno,
                        'is_async': isinstance(node, ast.AsyncFunctionDef)
                    }
                    file_info['functions'].append(func_info)
                
                # Import 분석
                elif isinstance(node, ast.Import):
                    for alias in node.names:
                        file_info['imports'].append(alias.name)
                        file_info['dependencies'].append(alias.name)
                
                elif isinstance(node, ast.ImportFrom):
                    if node.module:
                        file_info['imports'].append(node.module)
                        # 로컬 모듈 의존성 추적
                        if not node.module.startswith(('tkinter', 'pandas', 'gspread', 'matplotlib', 'selenium')):
                            file_info['dependencies'].append(node.module)
                
                # 상수 분석 (대문자 변수)
                elif isinstance(node, ast.Assign):
                    for target in node.targets:
                        if isinstance(target, ast.Name) and target.id.isupper():
                            file_info['constants'].append(target.id)
            
            return file_info
            
        except Exception as e:
            print(f"❌ 파일 분석 실패 {file_path}: {e}")
            return None
    
    def _get_name(self, node):
        """AST 노드에서 이름 추출"""
        if isinstance(node, ast.Name):
            return node.id
        elif isinstance(node, ast.Attribute):
            return f"{self._get_name(node.value)}.{node.attr}"
        else:
            return str(node)
    
    def _categorize_file(self, file_path):
        """파일 카테고리 분류"""
        name = file_path.stem.lower()
        
        if 'renewal' in name:
            if 'email' in name:
                return '📧 이메일 발송'
            elif 'quote' in name:
                return '📋 견적서 관리'
            elif 'search' in name:
                return '🔍 검색 기능'
            elif 'preview' in name or 'gui' in name:
                return '🖥️ 메인 UI'
            else:
                return '🔄 리뉴얼 관리'
        
        elif 'customer' in name:
            return '👥 고객사 관리'
        
        elif 'quote' in name:
            return '📋 견적서 관리'
        
        elif 'purchase' in name or 'sales' in name:
            return '💰 매입/매출 관리'
        
        elif 'dashboard' in name:
            if 'team' in name:
                return '📊 팀별 대시보드'
            elif 'foundry' in name:
                return '📊 Foundry 대시보드'
            elif 'interactive' in name:
                return '📊 인터랙티브 대시보드'
            else:
                return '📊 대시보드'
        
        elif 'license' in name or 'certificate' in name:
            return '🔐 인증서 관리'
        
        elif 'data_loader' in name:
            return '💾 데이터 로더'
        
        elif 'daou' in name:
            return '🔗 DAOU 연동'
        
        elif 'sendlog' in name or 'log' in name:
            return '📝 로그 관리'
        
        elif 'loading' in name:
            return '⏳ 로딩 UI'
        
        elif 'organization' in name:
            return '🏢 조직 관리'
        
        elif 'project' in name or 'version' in name:
            return '🛠️ 프로젝트 관리'
        
        else:
            return '🔧 기타 유틸리티'
    
    def _infer_purpose(self, file_path, tree):
        """파일 목적 추론"""
        name = file_path.stem.lower()
        docstring = ast.get_docstring(tree) or ''
        
        purposes = {
            'renewal_gui': '메인 애플리케이션 - 로그인, 탭 관리, 메인 UI 구성',
            'renewal_preview': '리뉴얼 데이터 조회, 필터링, 이메일/견적서 발송 미리보기',
            'renewal_email': 'Outlook을 통한 리뉴얼 안내 이메일 자동 발송',
            'renewal_email_multi': '다중 선택 이메일 발송 기능',
            'renewal_quote': 'Excel 견적서 자동 생성 및 Outlook 첨부',
            'renewal_quote_multi': '다중 견적서 생성 기능',
            'renewal_search': '리뉴얼 데이터 검색 및 필터링 유틸리티',
            'customer_info': '고객사 정보 조회, 관리, 거래 내역 추적',
            'customer_add': '신규 고객사 등록 다이얼로그',
            'customer_details': '고객사 상세 정보 표시',
            'customer_account_info': '고객사 계정 정보 관리',
            'customer_performance': '고객사별 성과 분석',
            'quote_info': '견적서 데이터 조회, Status 관리, 견적서 생성/수정',
            'quote_new': '신규 견적서 생성 다이얼로그',
            'quote_modify': '견적서 수정 다이얼로그',
            'quote_simple': '간편 견적서 생성',
            'quote_status_modify': '견적서 Status 일괄 수정',
            'quote_product_add': '견적서에 제품 추가',
            'purchase_sales_info': '매입/매출 데이터 조회, 기간별 필터링, 수익성 분석',
            'purchase_sales_modify': '매입/매출 데이터 수정',
            'purchase_sales_expdate_modify': '만료일 일괄 수정',
            'dashboard_info': '월별 영업사원별 통계 대시보드',
            'dashboard_foundry': 'Foundry 제품 전용 대시보드',
            'interactive_dashboard': '인터랙티브 차트 기반 종합 대시보드',
            'team_performance_dashboard': '팀별 성과 분석 대시보드 (Target 기반)',
            'license_certificate_info': '5개 제조사 인증서 생성 메인 UI',
            'license_certificate_BorisFX': 'BorisFX 인증서 자동 생성 (Word)',
            'license_certificate_Foundry': 'Foundry 인증서 및 라이선스 파일 생성 (Excel)',
            'license_certificate_Marmoset': 'Marmoset 인증서 생성 (Excel)',
            'license_certificate_Maxon': 'Maxon 인증서 생성 (Excel)',
            'license_certificate_VideoCopilot': 'VideoCopilot 인증서 생성 (Excel)',
            'data_loader': '구글 시트 데이터 로딩, 캐싱, 백그라운드 로딩 관리',
            'sendlog': '구글 시트 로그 기록 및 조회',
            'loading_indicator': '로딩 인디케이터 UI 표시',
            'loading_screen': '초기 로딩 스크린',
            'lisa_daou_sign': 'DAOU 전자결재 시스템 자동화 (Selenium)',
            'organization_manager': '조직 구조 관리 (부서/팀/영업사원)',
            'project_manager': '프로젝트 관리 기능',
            'resizable_treeview': '크기 조절 가능한 Treeview 위젯',
            'column_settings': '컬럼 설정 관리',
            'debug_target_columns': '디버깅용 컬럼 확인',
        }
        
        purpose = purposes.get(name, '')
        
        if not purpose and docstring:
            purpose = docstring.split('\n')[0][:100]
        
        if not purpose:
            purpose = f"{file_path.name} 모듈"
        
        return purpose
    
    def analyze_project(self):
        """프로젝트 전체 분석"""
        print("🔍 LISA 프로젝트 분석 시작...")
        
        # Python 파일 찾기
        py_files = []
        seen_files = set()  # 중복 방지
        
        # 제외할 디렉토리 및 파일 패턴
        exclude_dirs = ['__pycache__', 'build', 'dist', '.git', 'versioning', 'cache', '이전 파일', 'LISA 에 필요한 DB 모음. 엑셀 등']
        exclude_files = ['lisa_project_analyzer.py', 'import hashlib.py', 'login.py']
        exclude_patterns = ['복사본', 'test_', 'simple_', 'extracter', 'extraction', 'dashboard.py', 'daouoffice']
        
        # 루트 디렉토리의 Python 파일만 검색
        for py_file in self.project_root.glob('*.py'):
            # 이미 처리한 파일은 스킵
            if py_file.name in seen_files:
                continue
            
            # 제외할 파일 확인
            if py_file.name in exclude_files:
                continue
            
            # 제외 패턴 확인
            if any(pattern in py_file.name for pattern in exclude_patterns):
                continue
            
            py_files.append(py_file)
            seen_files.add(py_file.name)
        
        print(f"📁 발견된 Python 파일: {len(py_files)}개")
        
        # 각 파일 분석
        for py_file in sorted(py_files):
            print(f"   분석 중: {py_file.name}")
            result = self.analyze_file(py_file)
            if result:
                self.analysis_results.append(result)
        
        print(f"✅ 분석 완료: {len(self.analysis_results)}개 파일")
        return self.analysis_results
    
    def create_excel_report(self, output_file='LISA_프로젝트_기획서.xlsx'):
        """분석 결과를 Excel 기획서로 생성"""
        print(f"\n📊 Excel 기획서 생성 중: {output_file}")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 1. 프로젝트 개요 시트
            self._create_overview_sheet(writer)
            
            # 2. 파일 목록 시트
            self._create_file_list_sheet(writer)
            
            # 3. 모듈별 상세 시트
            self._create_module_details_sheet(writer)
            
            # 4. 클래스 목록 시트
            self._create_class_list_sheet(writer)
            
            # 5. 함수 목록 시트
            self._create_function_list_sheet(writer)
            
            # 6. 의존성 분석 시트
            self._create_dependency_sheet(writer)
            
            # 7. 카테고리별 분류 시트
            self._create_category_sheet(writer)
            
            # 8. 통계 시트
            self._create_statistics_sheet(writer)
        
        print(f"✅ Excel 기획서 생성 완료: {output_file}")
        return output_file
    
    def _create_overview_sheet(self, writer):
        """프로젝트 개요 시트"""
        overview_data = {
            '항목': [
                '프로젝트명',
                '버전',
                '생성일',
                '분석일',
                '총 Python 파일 수',
                '총 코드 라인 수',
                '총 클래스 수',
                '총 함수 수',
                '주요 기능',
                '기술 스택',
                '데이터베이스',
                '외부 연동',
                '패키징',
                '설명'
            ],
            '내용': [
                'LISA (License Information System for Administration)',
                'v2.01',
                '2024-2025',
                datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                len(self.analysis_results),
                sum(r['lines_of_code'] for r in self.analysis_results),
                sum(len(r['classes']) for r in self.analysis_results),
                sum(len(r['functions']) for r in self.analysis_results),
                'Renewal 관리, 고객사 관리, 견적서 관리, 매입/매출 관리, 대시보드, 인증서 관리',
                'Python 3.x, Tkinter, pandas, gspread, matplotlib, selenium, win32com',
                'Google Sheets (5개 워크시트: renewal_list, customer_list, quote_list, price_list, log_list)',
                'Google Sheets API, DAOU 전자결재, Outlook 이메일, Excel 자동화',
                'PyInstaller (단일 실행 파일)',
                '큐브렉스 소프트웨어 라이선스 관리 및 영업 지원 자동화 시스템'
            ]
        }
        
        df = pd.DataFrame(overview_data)
        df.to_excel(writer, sheet_name='📋 프로젝트 개요', index=False)
        
        # 열 너비 조정
        worksheet = writer.sheets['📋 프로젝트 개요']
        worksheet.column_dimensions['A'].width = 20
        worksheet.column_dimensions['B'].width = 80
    
    def _create_file_list_sheet(self, writer):
        """파일 목록 시트"""
        file_data = []
        for r in sorted(self.analysis_results, key=lambda x: (x['category'], x['file_name'])):
            file_data.append({
                '카테고리': r['category'],
                '파일명': r['file_name'],
                '모듈명': r['module_name'],
                '경로': r['file_path'],
                '목적': r['purpose'],
                '코드 라인수': r['lines_of_code'],
                '클래스 수': len(r['classes']),
                '함수 수': len(r['functions']),
                'Import 수': len(r['imports']),
                '상수 수': len(r['constants']),
                'Docstring': r['docstring'][:100] if r['docstring'] else ''
            })
        
        df = pd.DataFrame(file_data)
        df.to_excel(writer, sheet_name='📁 파일 목록', index=False)
        
        # 열 너비 조정
        worksheet = writer.sheets['📁 파일 목록']
        worksheet.column_dimensions['A'].width = 20
        worksheet.column_dimensions['B'].width = 30
        worksheet.column_dimensions['C'].width = 25
        worksheet.column_dimensions['D'].width = 40
        worksheet.column_dimensions['E'].width = 60
    
    def _create_module_details_sheet(self, writer):
        """모듈별 상세 시트"""
        module_data = []
        for r in sorted(self.analysis_results, key=lambda x: (x['category'], x['file_name'])):
            # 클래스 목록
            classes_str = '\n'.join([f"- {c['name']}: {c['docstring'][:50]}" for c in r['classes']]) if r['classes'] else '없음'
            
            # 함수 목록
            functions_str = '\n'.join([f"- {f['name']}({', '.join(f['args'])})" for f in r['functions'][:10]]) if r['functions'] else '없음'
            if len(r['functions']) > 10:
                functions_str += f"\n... 외 {len(r['functions']) - 10}개"
            
            # 주요 Import
            imports_str = ', '.join(sorted(set(r['imports']))[:15])
            if len(r['imports']) > 15:
                imports_str += f" ... 외 {len(r['imports']) - 15}개"
            
            module_data.append({
                '카테고리': r['category'],
                '모듈명': r['module_name'],
                '파일명': r['file_name'],
                '목적': r['purpose'],
                '코드 라인수': r['lines_of_code'],
                '클래스 목록': classes_str,
                '함수 목록': functions_str,
                'Import 목록': imports_str,
                '상수': ', '.join(r['constants'][:10]),
                'Docstring': r['docstring']
            })
        
        df = pd.DataFrame(module_data)
        df.to_excel(writer, sheet_name='🔍 모듈 상세', index=False)
        
        # 열 너비 조정
        worksheet = writer.sheets['🔍 모듈 상세']
        worksheet.column_dimensions['A'].width = 20
        worksheet.column_dimensions['B'].width = 25
        worksheet.column_dimensions['C'].width = 30
        worksheet.column_dimensions['D'].width = 60
        worksheet.column_dimensions['F'].width = 40
        worksheet.column_dimensions['G'].width = 40
    
    def _create_class_list_sheet(self, writer):
        """클래스 목록 시트"""
        class_data = []
        for r in self.analysis_results:
            for cls in r['classes']:
                class_data.append({
                    '카테고리': r['category'],
                    '파일명': r['file_name'],
                    '클래스명': cls['name'],
                    '설명': cls['docstring'][:200] if cls['docstring'] else '',
                    '상속': ', '.join(cls['bases']) if cls['bases'] else 'object',
                    '메서드 수': len(cls['methods']),
                    '주요 메서드': ', '.join(cls['methods'][:10]),
                    '라인번호': cls['line_no']
                })
        
        df = pd.DataFrame(class_data)
        df.to_excel(writer, sheet_name='🏛️ 클래스 목록', index=False)
        
        # 열 너비 조정
        worksheet = writer.sheets['🏛️ 클래스 목록']
        worksheet.column_dimensions['A'].width = 20
        worksheet.column_dimensions['B'].width = 30
        worksheet.column_dimensions['C'].width = 30
        worksheet.column_dimensions['D'].width = 60
        worksheet.column_dimensions['G'].width = 50
    
    def _create_function_list_sheet(self, writer):
        """함수 목록 시트"""
        func_data = []
        for r in self.analysis_results:
            for func in r['functions']:
                func_data.append({
                    '카테고리': r['category'],
                    '파일명': r['file_name'],
                    '함수명': func['name'],
                    '설명': func['docstring'][:200] if func['docstring'] else '',
                    '인자': ', '.join(func['args']),
                    '비동기': '예' if func['is_async'] else '아니오',
                    '라인번호': func['line_no']
                })
        
        df = pd.DataFrame(func_data)
        df.to_excel(writer, sheet_name='⚙️ 함수 목록', index=False)
        
        # 열 너비 조정
        worksheet = writer.sheets['⚙️ 함수 목록']
        worksheet.column_dimensions['A'].width = 20
        worksheet.column_dimensions['B'].width = 30
        worksheet.column_dimensions['C'].width = 30
        worksheet.column_dimensions['D'].width = 60
        worksheet.column_dimensions['E'].width = 40
    
    def _create_dependency_sheet(self, writer):
        """의존성 분석 시트"""
        dep_data = []
        for r in self.analysis_results:
            # 로컬 모듈 의존성만 추출
            local_deps = [d for d in r['dependencies'] if not d.startswith(('tkinter', 'pandas', 'gspread'))]
            
            if local_deps:
                dep_data.append({
                    '모듈명': r['module_name'],
                    '파일명': r['file_name'],
                    '카테고리': r['category'],
                    '의존 모듈 수': len(local_deps),
                    '의존 모듈': ', '.join(sorted(set(local_deps)))
                })
        
        df = pd.DataFrame(dep_data)
        df.to_excel(writer, sheet_name='🔗 모듈 의존성', index=False)
        
        # 열 너비 조정
        worksheet = writer.sheets['🔗 모듈 의존성']
        worksheet.column_dimensions['A'].width = 25
        worksheet.column_dimensions['B'].width = 30
        worksheet.column_dimensions['C'].width = 20
        worksheet.column_dimensions['E'].width = 60
    
    def _create_category_sheet(self, writer):
        """카테고리별 분류 시트"""
        category_stats = {}
        for r in self.analysis_results:
            cat = r['category']
            if cat not in category_stats:
                category_stats[cat] = {
                    '파일 수': 0,
                    '코드 라인 수': 0,
                    '클래스 수': 0,
                    '함수 수': 0,
                    '파일 목록': []
                }
            category_stats[cat]['파일 수'] += 1
            category_stats[cat]['코드 라인 수'] += r['lines_of_code']
            category_stats[cat]['클래스 수'] += len(r['classes'])
            category_stats[cat]['함수 수'] += len(r['functions'])
            category_stats[cat]['파일 목록'].append(r['file_name'])
        
        cat_data = []
        for cat, stats in sorted(category_stats.items()):
            cat_data.append({
                '카테고리': cat,
                '파일 수': stats['파일 수'],
                '코드 라인 수': stats['코드 라인 수'],
                '클래스 수': stats['클래스 수'],
                '함수 수': stats['함수 수'],
                '평균 파일 크기': stats['코드 라인 수'] // stats['파일 수'],
                '파일 목록': ', '.join(stats['파일 목록'])
            })
        
        df = pd.DataFrame(cat_data)
        df.to_excel(writer, sheet_name='📂 카테고리별 분류', index=False)
        
        # 열 너비 조정
        worksheet = writer.sheets['📂 카테고리별 분류']
        worksheet.column_dimensions['A'].width = 25
        worksheet.column_dimensions['G'].width = 60
    
    def _create_statistics_sheet(self, writer):
        """통계 시트"""
        total_files = len(self.analysis_results)
        total_lines = sum(r['lines_of_code'] for r in self.analysis_results)
        total_classes = sum(len(r['classes']) for r in self.analysis_results)
        total_functions = sum(len(r['functions']) for r in self.analysis_results)
        total_imports = sum(len(r['imports']) for r in self.analysis_results)
        
        # 가장 큰 파일
        largest_files = sorted(self.analysis_results, key=lambda x: x['lines_of_code'], reverse=True)[:10]
        
        # 가장 복잡한 파일 (클래스+함수 수)
        complex_files = sorted(self.analysis_results, 
                              key=lambda x: len(x['classes']) + len(x['functions']), 
                              reverse=True)[:10]
        
        stats_data = {
            '통계 항목': [
                '전체 파일 수',
                '전체 코드 라인 수',
                '평균 파일 크기',
                '전체 클래스 수',
                '파일당 평균 클래스 수',
                '전체 함수 수',
                '파일당 평균 함수 수',
                '전체 Import 수',
                '파일당 평균 Import 수',
                '',
                '가장 큰 파일 TOP 5',
                '', '', '', '', '',
                '가장 복잡한 파일 TOP 5',
                '', '', '', ''
            ],
            '값': [
                total_files,
                total_lines,
                f"{total_lines // total_files} 라인",
                total_classes,
                f"{total_classes / total_files:.1f}",
                total_functions,
                f"{total_functions / total_files:.1f}",
                total_imports,
                f"{total_imports / total_files:.1f}",
                '',
                f"1. {largest_files[0]['file_name']}: {largest_files[0]['lines_of_code']} 라인",
                f"2. {largest_files[1]['file_name']}: {largest_files[1]['lines_of_code']} 라인",
                f"3. {largest_files[2]['file_name']}: {largest_files[2]['lines_of_code']} 라인",
                f"4. {largest_files[3]['file_name']}: {largest_files[3]['lines_of_code']} 라인",
                f"5. {largest_files[4]['file_name']}: {largest_files[4]['lines_of_code']} 라인",
                '',
                f"1. {complex_files[0]['file_name']}: {len(complex_files[0]['classes'])+len(complex_files[0]['functions'])} 요소",
                f"2. {complex_files[1]['file_name']}: {len(complex_files[1]['classes'])+len(complex_files[1]['functions'])} 요소",
                f"3. {complex_files[2]['file_name']}: {len(complex_files[2]['classes'])+len(complex_files[2]['functions'])} 요소",
                f"4. {complex_files[3]['file_name']}: {len(complex_files[3]['classes'])+len(complex_files[3]['functions'])} 요소",
                f"5. {complex_files[4]['file_name']}: {len(complex_files[4]['classes'])+len(complex_files[4]['functions'])} 요소"
            ]
        }
        
        df = pd.DataFrame(stats_data)
        df.to_excel(writer, sheet_name='📈 통계', index=False)
        
        # 열 너비 조정
        worksheet = writer.sheets['📈 통계']
        worksheet.column_dimensions['A'].width = 30
        worksheet.column_dimensions['B'].width = 60


def main():
    """메인 실행 함수"""
    print("="*80)
    print("🚀 LISA 프로젝트 분석 및 Excel 기획서 생성 도구")
    print("="*80)
    print()
    
    # 프로젝트 분석
    analyzer = PythonFileAnalyzer(".")
    analyzer.analyze_project()
    
    # Excel 기획서 생성
    output_file = analyzer.create_excel_report()
    
    print()
    print("="*80)
    print(f"✅ 완료! Excel 파일이 생성되었습니다: {output_file}")
    print("="*80)
    
    return output_file


if __name__ == '__main__':
    main()

