import os
import subprocess
import shutil
from datetime import datetime

class VersionManager:
    def __init__(self, repo_root=None):
        self.repo_root = repo_root or os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
        self.version_file = os.path.join(self.repo_root, "versioning", "VERSION")
        self.changelog_file = os.path.join(self.repo_root, "versioning", "CHANGELOG.md")
        self.backup_dir = os.path.join(self.repo_root, "이전 파일")
        self.init_git_repo()

    def run_git(self, args):
        result = subprocess.run(
            ["git"] + args,
            cwd=self.repo_root,
            capture_output=True,
            text=True,
            encoding="utf-8"
        )
        if result.returncode != 0:
            raise Exception(f"Git 명령어 실패: {' '.join(args)}\n{result.stderr}")
        return result.stdout.strip()

    def init_git_repo(self):
        if not os.path.exists(os.path.join(self.repo_root, ".git")):
            self.run_git(["init"])
            self.run_git(["add", "."])
            self.run_git(["commit", "-m", "Initial commit"])

    def get_current_version(self):
        if os.path.exists(self.version_file):
            with open(self.version_file, "r", encoding="utf-8") as f:
                return f.read().strip()
        return "0.0.0"

    def set_version(self, version):
        with open(self.version_file, "w", encoding="utf-8") as f:
            f.write(version)

    def get_git_status(self):
        modified = self.run_git(["diff", "--name-only"]).split('\n')
        staged = self.run_git(["diff", "--cached", "--name-only"]).split('\n')
        return {
            "modified": [f for f in modified if f.strip()],
            "staged": [f for f in staged if f.strip()],
            "has_changes": bool(modified or staged)
        }

    def backup_python_files(self, version):
        """버전 태그 생성 시 .py 파일들을 날짜별 폴더에 백업"""
        try:
            # 현재 날짜와 시간으로 폴더명 생성
            now = datetime.now()
            date_folder = now.strftime("%Y-%m-%d_%H%M")
            backup_folder = os.path.join(self.backup_dir, f"{date_folder}_{version}")
            
            # 백업 폴더 생성
            os.makedirs(backup_folder, exist_ok=True)
            
            # 프로젝트 루트에서 모든 .py 파일 찾기
            python_files = []
            for root, dirs, files in os.walk(self.repo_root):
                # 이전 파일 폴더와 versioning 폴더는 제외
                if "이전 파일" in root or "versioning" in root:
                    continue
                for file in files:
                    if file.endswith('.py'):
                        python_files.append(os.path.join(root, file))
            
            # .py 파일들을 백업 폴더로 복사
            copied_files = []
            for py_file in python_files:
                # 상대 경로 계산
                rel_path = os.path.relpath(py_file, self.repo_root)
                # 백업 폴더 내에서 동일한 디렉토리 구조 유지
                backup_path = os.path.join(backup_folder, rel_path)
                backup_dir = os.path.dirname(backup_path)
                
                # 디렉토리가 없으면 생성
                if not os.path.exists(backup_dir):
                    os.makedirs(backup_dir, exist_ok=True)
                
                # 파일 복사
                shutil.copy2(py_file, backup_path)
                copied_files.append(rel_path)
            
            # 백업 정보 파일 생성
            backup_info = {
                "version": version,
                "backup_date": now.isoformat(),
                "backup_folder": backup_folder,
                "copied_files": copied_files,
                "total_files": len(copied_files)
            }
            
            info_file = os.path.join(backup_folder, "backup_info.txt")
            with open(info_file, "w", encoding="utf-8") as f:
                f.write(f"버전: {version}\n")
                f.write(f"백업 날짜: {now.strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"백업 폴더: {backup_folder}\n")
                f.write(f"총 파일 수: {len(copied_files)}\n")
                f.write(f"\n백업된 파일 목록:\n")
                for file_path in copied_files:
                    f.write(f"- {file_path}\n")
            
            return backup_info
            
        except Exception as e:
            raise Exception(f"파일 백업 중 오류 발생: {e}")

    def create_version_tag(self, version, message=""):
        # 버전 태그 생성 전에 .py 파일들 백업
        backup_info = self.backup_python_files(version)
        
        self.set_version(version)
        self.run_git(["add", "versioning/VERSION"])
        self.run_git(["commit", "-m", f"Bump version to {version}"])
        self.run_git(["tag", "-a", f"v{version}", "-m", message or f"Version {version}"])
        
        return backup_info

    def get_version_history(self):
        tags = self.run_git(["tag", "--sort=-version:refname"]).split('\n')
        history = []
        for tag in tags:
            if not tag.strip():
                continue
            commit_hash = self.run_git(["rev-list", "-n", "1", tag])
            commit_date = self.run_git(["log", "-1", "--format=%cd", "--date=iso", tag])
            commit_message = self.run_git(["log", "-1", "--format=%s", tag])
            history.append({
                "version": tag,
                "commit_hash": commit_hash,
                "date": commit_date,
                "message": commit_message
            })
        return history

    def get_diff_between_versions(self, old_version, new_version):
        """두 버전(태그) 사이의 git diff를 반환합니다."""
        return self.run_git(["diff", f"{old_version}", f"{new_version}"])

    def checkout_version(self, version):
        """특정 태그(버전)로 체크아웃합니다."""
        return self.run_git(["checkout", version])

    def get_backup_list(self):
        """백업된 버전 목록을 반환합니다."""
        if not os.path.exists(self.backup_dir):
            return []
        
        backups = []
        for item in os.listdir(self.backup_dir):
            item_path = os.path.join(self.backup_dir, item)
            if os.path.isdir(item_path):
                # 폴더명에서 날짜와 버전 추출 (YYYY-MM-DD_HHMM_version 형식)
                try:
                    parts = item.split('_')
                    if len(parts) >= 3:
                        date_part = f"{parts[0]}_{parts[1]}_{parts[2]}"
                        version_part = '_'.join(parts[3:]) if len(parts) > 3 else "unknown"
                        
                        info_file = os.path.join(item_path, "backup_info.txt")
                        if os.path.exists(info_file):
                            with open(info_file, "r", encoding="utf-8") as f:
                                info_content = f.read()
                        else:
                            info_content = "정보 파일 없음"
                        
                        backups.append({
                            "folder": item,
                            "date": date_part,
                            "version": version_part,
                            "path": item_path,
                            "info": info_content
                        })
                except:
                    continue
        
        # 날짜순으로 정렬 (최신순)
        backups.sort(key=lambda x: x["date"], reverse=True)
        return backups

version_manager = VersionManager() 