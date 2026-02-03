# project_manager.py - 프로젝트 저장 관리
import os
import shutil
import zipfile
import json
from datetime import datetime
from config import config

class ProjectManager:
    """프로젝트 저장 및 백업 관리"""
    
    def __init__(self, project_root="."):
        self.project_root = os.path.abspath(project_root)
        self.backup_folder = os.path.join(self.project_root, "backups")
        self.ensure_backup_folder()
    
    def ensure_backup_folder(self):
        """백업 폴더 생성"""
        if not os.path.exists(self.backup_folder):
            os.makedirs(self.backup_folder)
    
    def save_project_state(self, state_name=None):
        """현재 프로젝트 상태 저장"""
        if state_name is None:
            state_name = f"project_state_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        # 설정 업데이트
        config.update_config("last_save", datetime.now().isoformat())
        config.update_config("current_state", state_name)
        
        # 상태 정보 저장
        state_info = {
            "state_name": state_name,
            "timestamp": datetime.now().isoformat(),
            "project_root": self.project_root,
            "config": config.config
        }
        
        state_file = os.path.join(self.backup_folder, f"{state_name}.json")
        with open(state_file, 'w', encoding='utf-8') as f:
            json.dump(state_info, f, ensure_ascii=False, indent=2)
        
        return state_file
    
    def create_project_backup(self, backup_name=None):
        """프로젝트 전체 백업 생성"""
        if backup_name is None:
            backup_name = f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        backup_path = os.path.join(self.backup_folder, f"{backup_name}.zip")
        
        # 제외할 파일/폴더 목록
        exclude_patterns = [
            '__pycache__',
            '.git',
            'backups',
            '*.pyc',
            '*.log',
            'temp',
            'tmp'
        ]
        
        with zipfile.ZipFile(backup_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(self.project_root):
                # 제외할 폴더 제거
                dirs[:] = [d for d in dirs if not any(pattern in d for pattern in exclude_patterns)]
                
                for file in files:
                    # 제외할 파일 건너뛰기
                    if any(pattern in file for pattern in exclude_patterns):
                        continue
                    
                    file_path = os.path.join(root, file)
                    arc_name = os.path.relpath(file_path, self.project_root)
                    zipf.write(file_path, arc_name)
        
        # 백업 정보 저장
        backup_info = {
            "backup_name": backup_name,
            "timestamp": datetime.now().isoformat(),
            "backup_path": backup_path,
            "file_size": os.path.getsize(backup_path)
        }
        
        info_file = os.path.join(self.backup_folder, f"{backup_name}_info.json")
        with open(info_file, 'w', encoding='utf-8') as f:
            json.dump(backup_info, f, ensure_ascii=False, indent=2)
        
        return backup_path
    
    def restore_project_backup(self, backup_name):
        """백업에서 프로젝트 복원"""
        backup_path = os.path.join(self.backup_folder, f"{backup_name}.zip")
        
        if not os.path.exists(backup_path):
            raise FileNotFoundError(f"백업 파일을 찾을 수 없습니다: {backup_path}")
        
        # 임시 복원 폴더 생성
        temp_restore = os.path.join(self.backup_folder, "temp_restore")
        if os.path.exists(temp_restore):
            shutil.rmtree(temp_restore)
        os.makedirs(temp_restore)
        
        # 백업 압축 해제
        with zipfile.ZipFile(backup_path, 'r') as zipf:
            zipf.extractall(temp_restore)
        
        # 현재 프로젝트 백업
        current_backup = self.create_project_backup("before_restore")
        
        # 파일 복원
        for item in os.listdir(temp_restore):
            src = os.path.join(temp_restore, item)
            dst = os.path.join(self.project_root, item)
            
            if os.path.exists(dst):
                if os.path.isdir(dst):
                    shutil.rmtree(dst)
                else:
                    os.remove(dst)
            
            if os.path.isdir(src):
                shutil.copytree(src, dst)
            else:
                shutil.copy2(src, dst)
        
        # 임시 폴더 정리
        shutil.rmtree(temp_restore)
        
        return True
    
    def list_backups(self):
        """백업 목록 조회"""
        backups = []
        for file in os.listdir(self.backup_folder):
            if file.endswith('.zip'):
                backup_name = file[:-4]  # .zip 제거
                info_file = os.path.join(self.backup_folder, f"{backup_name}_info.json")
                
                backup_info = {
                    "name": backup_name,
                    "file": file,
                    "timestamp": None,
                    "size": os.path.getsize(os.path.join(self.backup_folder, file))
                }
                
                if os.path.exists(info_file):
                    try:
                        with open(info_file, 'r', encoding='utf-8') as f:
                            info = json.load(f)
                            backup_info["timestamp"] = info.get("timestamp")
                    except:
                        pass
                
                backups.append(backup_info)
        
        return sorted(backups, key=lambda x: x["timestamp"] or "", reverse=True)
    
    def cleanup_old_backups(self, keep_count=10):
        """오래된 백업 정리"""
        backups = self.list_backups()
        if len(backups) <= keep_count:
            return
        
        # 오래된 백업 삭제
        for backup in backups[keep_count:]:
            backup_file = os.path.join(self.backup_folder, backup["file"])
            info_file = os.path.join(self.backup_folder, f"{backup['name']}_info.json")
            
            if os.path.exists(backup_file):
                os.remove(backup_file)
            if os.path.exists(info_file):
                os.remove(info_file)

# 전역 프로젝트 매니저 인스턴스
project_manager = ProjectManager() 