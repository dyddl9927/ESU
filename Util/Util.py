import winreg
import socket
import os,sys
import ctypes
from tkinter import messagebox

class Util:
    def __init__(self, root=None):
        self.root = root
    @staticmethod
    def get_windows_build():
        """Windows 빌드 번호 정확하게 가져오기 (레지스트리 사용)"""
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                                r"SOFTWARE\Microsoft\Windows NT\CurrentVersion")
            
            # CurrentBuild (예: 19045)
            current_build, _ = winreg.QueryValueEx(key, "CurrentBuild")
            
            # UBR (Update Build Revision, 예: 6456)
            try:
                ubr, _ = winreg.QueryValueEx(key, "UBR")
                full_build = f"{current_build}.{ubr}"
            except:
                full_build = current_build
            
            winreg.CloseKey(key)
            return full_build
            
        except:
            return "알 수 없음"
    
    @staticmethod
    def check_windows_version():
        """Windows 10 버전이 19045.6456인지 확인"""
        required_version = "19045.6456"
        current_build = Util.get_windows_build()
        
        return current_build == required_version
    
    @staticmethod
    def get_local_ip():
        """현재 PC의 IP 주소 가져오기"""
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            ip = s.getsockname()[0]
            s.close()
            return ip
        except:
            return "127.0.0.1"
        
    @staticmethod
    def _default_icon_path():
        repo_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        candidates = [
            os.path.join(repo_root, "img", "logo2.ico"),
            os.path.join(repo_root, "logo2.ico"),
        ]
        for path in candidates:
            if os.path.exists(path):
                return path
        return ""
    
    def is_admin(self):
        """관리자 권한 확인"""
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False
    
    def request_admin(self):
        """관리자 권한으로 재실행"""
        result = messagebox.askokcancel(
                "관리자 권한 필요",
                "이 프로그램은 관리자 권한이 필요합니다.\n관리자 권한으로 다시 시작하시겠습니까?"
        )
        if result:
            try:
                ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", sys.executable, " ".join(sys.argv), None, 1
                )
            except:
                print("관리자 권한 요청 실패")
                messagebox.showerror("오류", "관리자 권한으로 실행할 수 없습니다.")
        self.root.quit()