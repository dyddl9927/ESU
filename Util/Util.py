import winreg
import socket
import os,sys
import ctypes
from tkinter import messagebox

class Util:

    def __init__(self, root=None):
        self.root = root

    # 윈도우 빌드 번호 가져오기(레지스트리 값)
    @staticmethod
    def get_windows_build():
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
    
    # 윈도우 10 버전이 19045.6456인지 확인
    @staticmethod
    def check_windows_version():
        flag = False
        required_version = "19045.6456"
        current_build = Util.get_windows_build()

        if current_build == "19045.6456" or current_build == "19045.6466":
            flag = True
        
        return flag
    
    # 현재 PC의 IP 주소 가져오기
    @staticmethod
    def get_local_ip():
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            ip = s.getsockname()[0]
            s.close()
            return ip
        except:
            return "127.0.0.1"
    
    # 기본 아이콘 경로 반환
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
    
    # 관리자 권한 확인  
    def is_admin(self):
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False
    
    # 관리자 권한으로 재실행
    def request_admin(self):
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