import winreg
import socket
import os,sys
import ctypes
import subprocess
import re
from datetime import datetime
from pathlib import Path
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

    @staticmethod
    def get_log_path():
        try:
            log_dir = Path(r"C:\ESU")
            log_dir.mkdir(parents=True, exist_ok=True)
            return log_dir / "esu_log.txt"
        except Exception:
            repo_root = Path(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            return repo_root / "esu_log.txt"

    @staticmethod
    def write_log(message, details=None):
        try:
            log_path = Util.get_log_path()
            lines = [
                "=" * 80,
                f"시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                f"내용: {message}",
            ]
            if details:
                lines.append("상세:")
                lines.extend(str(details).splitlines())
            lines.append("")

            with open(log_path, "a", encoding="utf-8") as log_file:
                log_file.write("\n".join(lines))
        except Exception:
            pass

    @staticmethod
    def write_ip_error_log(message, details=None):
        Util.write_log(message, details)

    @staticmethod
    def is_valid_local_ipv4(ip):
        return (
            ip
            and not ip.startswith("127.")
            and not ip.startswith("169.254.")
            and ip != "0.0.0.0"
        )

    @staticmethod
    def get_ipconfig_output():
        try:
            result = subprocess.run(
                "ipconfig /all",
                shell=True,
                capture_output=True,
                text=True,
                encoding="cp949",
                errors="ignore",
                timeout=10,
            )
            return result.stdout + result.stderr
        except Exception as e:
            return f"ipconfig 실행 실패: {e}"

    @staticmethod
    def extract_ipv4_from_ipconfig(output):
        ips = []
        for line in output.splitlines():
            if "IPv4" not in line:
                continue
            match = re.search(r"(\d{1,3}(?:\.\d{1,3}){3})", line)
            if match and Util.is_valid_local_ipv4(match.group(1)):
                ips.append(match.group(1))
        return ips
    
    # 현재 PC의 IP 주소 가져오기
    @staticmethod
    def get_local_ip():
        errors = []

        try:
            with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
                s.settimeout(3)
                s.connect(("8.8.8.8", 80))
                ip = s.getsockname()[0]
                if Util.is_valid_local_ipv4(ip):
                    Util.write_log(
                        "IP 조회 성공",
                        f"method: udp_route_8.8.8.8\nip: {ip}",
                    )
                    return ip
                errors.append(f"8.8.8.8 기준 IP가 유효하지 않음: {ip}")
        except Exception as e:
            errors.append(f"8.8.8.8 기준 IP 확인 실패: {type(e).__name__}: {e}")

        try:
            hostname = socket.gethostname()
            candidate_ips = socket.gethostbyname_ex(hostname)[2]
            for ip in candidate_ips:
                if Util.is_valid_local_ipv4(ip):
                    Util.write_log(
                        "기본 방식으로 IP를 가져오지 못해 호스트명 기준 IP를 사용했습니다.",
                        "\n".join(errors + [f"hostname: {hostname}", f"candidate_ips: {candidate_ips}"]),
                    )
                    return ip
            errors.append(f"호스트명 기준 유효 IP 없음. hostname: {hostname}, candidate_ips: {candidate_ips}")
        except Exception as e:
            errors.append(f"호스트명 기준 IP 확인 실패: {type(e).__name__}: {e}")

        ipconfig_output = Util.get_ipconfig_output()
        ipconfig_ips = Util.extract_ipv4_from_ipconfig(ipconfig_output)
        if ipconfig_ips:
            Util.write_log(
                "기본 방식과 호스트명 방식으로 IP를 가져오지 못해 ipconfig 기준 IP를 사용했습니다.",
                "\n".join(errors + [f"ipconfig_ips: {ipconfig_ips}"]),
            )
            return ipconfig_ips[0]

        Util.write_log(
            "현재 PC의 유효한 IPv4 주소를 가져오지 못해 127.0.0.1로 처리했습니다.",
            "\n".join(errors + ["", "[ipconfig /all]", ipconfig_output]),
        )
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
