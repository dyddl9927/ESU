from tkinter import messagebox
import subprocess
import openpyxl
from openpyxl import Workbook
import os
from pathlib import Path
import re
import json
import threading
import queue
from concurrent.futures import Future

class Service :
    def __init__(self, root, ui, util):
        self.root = root
        self.ui = ui
        self.util = util
        self.busy = False
        self.ui_thread = threading.get_ident()
        self.ui_queue = queue.Queue()

        # 엑셀파일 저장경로

        #self.excel_path = r"\\172.26.21.20\ESU\esu.xlsx" #인천공항 공유
        #self.excel_path = r"\\172.28.243.228\ESU\esu.xlsx" #내부테스트용
        self.excel_path = r"c:\ESU\esu.xlsx"

        # 인증 cdkey 값 입력
        self.cdkey = "PGD98-N4MWG-9YKMW-P83F4-PPVPJ"

        # 정품 인증 ID값 입력
        self.activation_id = "f520e45e-7413-4a34-a497-d2765967d094"
        self.current_ip = self.util.get_local_ip()
        self.current_version = self.util.get_windows_build()

    def ensure_valid_ip(self):
        if self.util.is_valid_local_ipv4(self.current_ip):
            return True

        log_path = self.util.get_log_path()
        self.show_message("showerror",
            "IP 확인 오류",
            f"현재 PC의 유효한 IP를 가져오지 못했습니다.\n\n"
            f"현재 IP: {self.current_ip}\n"
            f"로그 파일: {log_path}\n\n"
            f"네트워크 연결, 방화벽, VPN, IP 할당 상태를 확인해주세요."
        )
        self.set_status(text="IP 확인 실패", fg="red")
        return False

    def start_task(self, task):
        if self.busy:
            return
        self.busy = True
        self.ui.btn_auth.config(state="disabled")
        self.root.after(50, self.poll_ui)

        def worker():
            try:
                task()
            except Exception as e:
                self.util.write_log("작업 실패", f"{type(e).__name__}: {e}")
                self.set_status(text="작업 실패", fg="red")
                self.show_message("showerror", "작업 오류", str(e))
            finally:
                self.ui_queue.put((self.finish_task, (), {}, None))

        threading.Thread(target=worker, daemon=True).start()

    def finish_task(self):
        self.busy = False
        self.ui.btn_auth.config(state="normal")

    def poll_ui(self):
        try:
            while True:
                callback, args, kwargs, future = self.ui_queue.get_nowait()
                try:
                    result = callback(*args, **kwargs)
                    if future is not None:
                        future.set_result(result)
                except Exception as e:
                    if future is not None:
                        future.set_exception(e)
        except queue.Empty:
            pass
        if self.busy:
            self.root.after(50, self.poll_ui)

    def call_ui(self, callback, *args, **kwargs):
        if threading.get_ident() == self.ui_thread:
            return callback(*args, **kwargs)
        future = Future()
        self.ui_queue.put((callback, args, kwargs, future))
        return future.result()

    def set_status(self, **kwargs):
        self.call_ui(self.ui.status_label.config, **kwargs)

    def show_message(self, kind, *args):
        return self.call_ui(getattr(messagebox, kind), *args)

    def initialize(self):
        self.check_installation_id()

    def prepare_operation(self):
        if not self.ensure_valid_ip():
            return False
        if not self.ensure_excel_exists():
            self.set_status(text="엑셀 준비 실패 - 작업 중단", fg="red")
            return False
        return True

    def ensure_excel_exists(self):
        wb = None
        try:
            Path(self.excel_path).parent.mkdir(parents=True, exist_ok=True)
            if os.path.exists(self.excel_path):
                wb = openpyxl.load_workbook(self.excel_path)
                ws = wb.active
                ws['D1'] = 'END'
                ws['E1'] = 'PRE'
            else:
                wb = Workbook()
                ws = wb.active
                ws.title = "ESU인증"
                for column, header in enumerate(("IP", "DTI", "확인", "END", "PRE"), 1):
                    ws.cell(1, column, header)
                for column, width in (("A", 20), ("B", 70), ("C", 70)):
                    ws.column_dimensions[column].width = width
            wb.save(self.excel_path)
            return True
        except Exception as e:
            self.util.write_log("엑셀 준비 실패", f"{type(e).__name__}: {e}")
            self.show_message("showerror", "파일 오류", "엑셀 파일을 준비하지 못했습니다. 파일이 열려 있는지, 저장 권한과 파일 상태를 확인해주세요.")
            return False
        finally:
            if wb is not None:
                wb.close()

    def run_command(self, command, timeout=60):
        result = subprocess.run(
            command, shell=False, capture_output=True, text=True,
            encoding="cp949", errors="replace", timeout=timeout,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        output = result.stdout + result.stderr
        if result.returncode != 0 or re.search(r"0x[0-9a-fA-F]{8}", output):
            raise RuntimeError(output.strip() or "명령 실행 실패")
        return output

    def slmgr_command(self, *args):
        system32 = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32"
        return [str(system32 / "cscript.exe"), "//Nologo", str(system32 / "slmgr.vbs"), *args]

    def query_esu_license(self):
        if not re.fullmatch(r"[0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}", self.activation_id):
            raise ValueError("ESU 정품 인증 ID가 올바르지 않습니다.")
        script = (
            "$ErrorActionPreference = 'Stop'; "
            f"$license = Get-CimInstance -ClassName SoftwareLicensingProduct -Filter \"ID='{self.activation_id}'\"; "
            "if ($null -eq $license) { 'null' } else { "
            "$license | Select-Object LicenseStatus,PartialProductKey | ConvertTo-Json -Compress }"
        )
        output = self.run_command(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script], timeout=30,
        )
        license = json.loads(output)
        if license is None:
            return None
        if not isinstance(license, dict) or type(license.get("LicenseStatus")) is not int or license["LicenseStatus"] not in range(7):
            raise ValueError("ESU 라이선스 조회 결과가 올바르지 않습니다.")
        return license

    def check_existing_esu(self):
        """설정된 ESU 인증 ID가 사용 중이거나 조회 실패이면 작업을 중단한다."""
        self.set_status(text="기존 ESU 사용 여부 확인 중...", fg="blue")
        try:
            license = self.query_esu_license()
            status = license["LicenseStatus"] if license is not None else -1
        except Exception as e:
            self.util.write_log(
                "기존 ESU 사용 여부 조회 실패",
                f"ip: {self.current_ip}\nactivation_id: {self.activation_id}\nerror: {type(e).__name__}: {e}",
            )
            self.set_status(text="ESU 상태 조회 실패 - 작업 중단", fg="red")
            self.show_message("showerror",
                "ESU 상태 확인 오류",
                "기존 ESU 사용 여부를 확인하지 못해 작업을 중단했습니다.\n"
                "관리자 권한 및 Windows 라이선스 서비스를 확인한 후 다시 시도해주세요.",
            )
            return True

        if status != 1:
            return False

        saved = self.save_license_status_to_excel("사용 허가됨")
        self.util.write_log(
            "이미 인증된 ESU: 키 설치 및 인증 생략",
            f"ip: {self.current_ip}\nactivation_id: {self.activation_id}",
        )
        self.set_status(
            text="ESU 사용 중 - 작업 생략" if saved else "ESU 사용 중 - 상태 저장 실패",
            fg="green" if saved else "orange",
        )
        self.show_message("showinfo",
            "ESU 사용 중",
            "이 PC는 이미 ESU가 인증되어 사용 중입니다.\nCD 키 설치 및 재인증을 생략합니다.\n\n"
            + ("엑셀 END 열에 '사용 허가됨'으로 기록했습니다." if saved else "엑셀에 상태를 저장하지 못했습니다."),
        )
        return True

    # A작업: 기존 ESU 확인 후 Windows 버전 확인, CD키 설치 및 설치ID 수집
    def check_installation_id(self):
        if not self.prepare_operation():
            return

        if self.check_existing_esu():
            return

        # 먼저 Windows 버전 확인
        if not self.util.check_windows_version():
            # 버전이 맞지 않으면 PRE 열에 "설치 불가" 저장
            saved = self.save_pre_status("설치 불가")
            self.util.write_log(
                "Windows 버전 불일치로 설치 불가 처리",
                f"ip: {self.current_ip}\ncurrent_version: {self.current_version}\nrequired_version: 19045.6456 또는 19045.6466",
            )
            self.show_message("showerror",
                "버전 불일치",
                f"Windows 10 버전이 요구사항과 맞지 않습니다.\n\n"
                f"현재 버전: {self.current_version}\n"
                f"필요 버전: 19045.6456 또는 19045.6466\n\n"
                + ("엑셀 PRE 열에 '설치 불가'로 기록되었습니다." if saved else "엑셀 PRE 상태 저장에 실패했습니다.")
            )
            self.set_status(text="버전 불일치 - 설치 불가", fg="red")
            return

        # 버전이 맞으면 CD 키 확인
        if not self.cdkey:
            self.show_message("showwarning", "입력 오류", "CD 키를 입력해주세요!")
            return

        self.set_status(text="설치ID 수집 중...", fg="blue")

        try:
            if not re.fullmatch(r"[A-Za-z0-9]{5}(?:-[A-Za-z0-9]{5}){4}", self.cdkey):
                raise ValueError("CD 키 형식이 올바르지 않습니다.")
            self.run_command(self.slmgr_command("/ipk", self.cdkey))
            license = self.query_esu_license()
            if license is None or (license.get("PartialProductKey") or "").upper() != self.cdkey[-5:].upper():
                raise RuntimeError("설정된 ESU 제품에 CD 키가 등록된 것을 확인하지 못했습니다.")
            result_dti = self.run_command(self.slmgr_command("/dti", self.activation_id))
        except Exception as e:
            self.util.write_log("CD 키 설치 또는 설치 ID 조회 실패", f"{type(e).__name__}: {e}")
            self.show_message("showerror", "설치 오류", str(e))
            self.set_status(text="설치 실패 - 작업 중단", fg="red")
            return

        dti_value = self.extract_dti_value(result_dti)

        if dti_value:
            # 엑셀 저장 시도 (성공/실패 반환 받음)
            if self.save_to_excel(dti_value):
                self.show_message("showinfo", "성공", f"설치ID가 수집되었습니다.\n\nDTI: {dti_value}\n\n엑셀 파일에 저장되었습니다.")
                self.set_status(text="설치 완료", fg="green")
            else:
                self.set_status(text="엑셀 저장 실패", fg="red")
        else:
            self.show_message("showerror", "오류", f"설치ID 수집 실패\n\n{result_dti}")
            self.set_status(text="설치 실패", fg="red")

    # PRE 열에 상태 저장(윈도우 버전 불일치 시 사용)
    def save_pre_status(self, status):
        wb = None
        try:
            wb = openpyxl.load_workbook(self.excel_path)
            ws = wb.active

            ip_found = False
            for row in range(2, ws.max_row + 1):
                if ws.cell(row, 1).value == self.current_ip:
                    ws.cell(row, 5, status)  # E열(5번째 열)에 저장
                    ip_found = True
                    break

            if not ip_found:
                new_row = ws.max_row + 1
                ws.cell(new_row, 1, self.current_ip)
                ws.cell(new_row, 5, status)

            wb.save(self.excel_path)
            return True

        except PermissionError:
            self.util.write_log(
                "PRE 상태 저장 실패: 엑셀 파일 열림",
                f"ip: {self.current_ip}\nstatus: {status}\nexcel_path: {self.excel_path}",
            )
            self.show_message("showerror", "저장 오류", "엑셀 파일(esu.xlsx)이 열려있습니다.\n파일을 닫고 다시 시도해주세요.")
            return False
        except Exception as e:
            self.util.write_log(
                "PRE 상태 저장 중 예외 발생",
                f"ip: {self.current_ip}\nstatus: {status}\nexcel_path: {self.excel_path}\nerror: {type(e).__name__}: {e}",
            )
            self.show_message("showerror", "오류", f"엑셀 저장 중 오류 발생: {e}")
            return False
        finally:
            if wb is not None:
                wb.close()

    # 설치ID(DTI) 값 추출

    def extract_dti_value(self, output):
        for line in output.splitlines():
            match = re.match(r"\s*(?:설치\s*ID|Installation\s*ID)\s*:\s*(.*)$", line, re.IGNORECASE)
            candidate = match.group(1) if match else line.strip()
            if re.fullmatch(r"[0-9\s-]+", candidate):
                digits = re.sub(r"[\s-]", "", candidate)
                if len(digits) in (54, 63):
                    return digits
        return None

    # 엑셀에 IP와 DTI 저장
    def save_to_excel(self, dti_value):
        wb = None
        try:
            wb = openpyxl.load_workbook(self.excel_path)
            ws = wb.active

            ip_found = False
            for row in range(2, ws.max_row + 1):
                if ws.cell(row, 1).value == self.current_ip:
                    ws.cell(row, 2, dti_value)
                    ip_found = True
                    break

            if not ip_found:
                new_row = ws.max_row + 1
                ws.cell(new_row, 1, self.current_ip)
                ws.cell(new_row, 2, dti_value)

            wb.save(self.excel_path)
            return True

        except PermissionError:
            self.util.write_log(
                "DTI 저장 실패: 엑셀 파일 열림",
                f"ip: {self.current_ip}\nexcel_path: {self.excel_path}",
            )
            self.show_message("showerror", "저장 오류", "엑셀 파일(esu.xlsx)이 열려있습니다.\n파일을 닫고 다시 시도해주세요.")
            return False
        except Exception as e:
            self.util.write_log(
                "DTI 저장 중 예외 발생",
                f"ip: {self.current_ip}\nexcel_path: {self.excel_path}\nerror: {type(e).__name__}: {e}",
            )
            self.show_message("showerror", "오류", f"엑셀 저장 중 오류 발생: {e}")
            return False
        finally:
            if wb is not None:
                wb.close()

    # B작업: 인증 및 라이선스 상태 확인

    def activate_esu(self):
        if not self.prepare_operation():
            return

        if self.check_existing_esu():
            return

        self.set_status(text="인증 진행 중...", fg="blue")

        # 엑셀에서 확인 값 가져오기
        confirm_value = self.get_confirm_value_from_excel()

        if confirm_value == "OPEN_ERROR":
            self.set_status(text="엑셀 파일 열려있음", fg="red")
            return

        if not confirm_value:
            self.show_message("showerror", "오류", f"엑셀 파일에서 IP({self.current_ip})에 해당하는 확인 값을 찾을 수 없습니다.\n'확인' 항목에 값을 입력해주세요.")
            self.set_status(text="인증 실패 (값 없음)", fg="red")
            return

        confirm_value = re.sub(r"[\s-]", "", confirm_value)
        if not re.fullmatch(r"[0-9]{48}", confirm_value):
            self.show_message("showerror", "확인 값 오류", "확인 ID는 48자리 숫자여야 합니다. 엑셀 확인 열을 텍스트 형식으로 입력해주세요.")
            self.set_status(text="인증 실패 (확인 ID 형식)", fg="red")
            return
        try:
            result_atp = self.run_command(self.slmgr_command("/atp", confirm_value, self.activation_id))
            self.set_status(text="라이선스 상태 확인 중...", fg="blue")
            license = self.query_esu_license()
        except Exception as e:
            self.util.write_log("ESU 인증 또는 상태 조회 실패", f"{type(e).__name__}: {e}")
            self.show_message("showerror", "인증 오류", str(e))
            self.set_status(text="인증 실패 - 작업 중단", fg="red")
            return

        license_status = "사용 허가됨"
        if license is not None and license["LicenseStatus"] == 1:
            # 엑셀에 라이선스 상태 저장
            if self.save_license_status_to_excel(license_status):
                self.show_message("showinfo", "성공",
                    f"인증이 완료되었습니다!\n\n"
                    f"인증 결과: 인증성공!\n\n"
                    f"라이선스 상태: {license_status}\n\n"
                    f"엑셀 파일 END 열에 저장되었습니다.")
                self.set_status(text="인증 및 상태 확인 완료", fg="green")
            else:
                self.set_status(text="엑셀 저장 실패", fg="red")
        else:
            self.show_message("showwarning", "결과",
                f"인증 결과:\n{result_atp}\n\n"
                f"설정된 ESU 제품이 사용 허가됨 상태가 아닙니다.")
            self.set_status(text="인증 실패 (상태 미확인)", fg="orange")

    # ESU 라이선스 상태 추출
    def extract_license_status(self, dlv_output):
        block = self.extract_esu_block(dlv_output)
        if block:
            status = self.extract_license_status_from_block(block)
            if status:
                return status
        return self.extract_license_status_from_block(dlv_output)

    def extract_esu_activation_id(self, dlv_output):
        block = self.extract_esu_block(dlv_output)
        if not block:
            return None
        pattern_kor = r"정품 인증 ID\s*:\s*([^\r\n]+)"
        match = re.search(pattern_kor, block, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        pattern_eng = r"Activation ID\s*:\s*([^\r\n]+)"
        match_eng = re.search(pattern_eng, block, re.IGNORECASE)
        if match_eng:
            return match_eng.group(1).strip()
        return None

    def extract_esu_block(self, dlv_output):
        blocks = re.split(r"\r?\n\r?\n+", dlv_output)
        for block in blocks:
            if re.search(r"\bESU\b", block, re.IGNORECASE):
                return block
        for block in blocks:
            if re.search(r"부분 제품 키\s*:\s*QDFWW", block, re.IGNORECASE):
                return block
            if re.search(r"Partial Product Key\s*:\s*QDFWW", block, re.IGNORECASE):
                return block
        return None

    def extract_license_status_from_block(self, text):
        pattern_kor = r"라이선스 상태\s*:\s*([^\r\n]+)"
        match = re.search(pattern_kor, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        pattern_eng = r"License Status\s*:\s*([^\r\n]+)"
        match_eng = re.search(pattern_eng, text, re.IGNORECASE)
        if match_eng:
            return match_eng.group(1).strip()
        return None

    # 현재 PC IP에 해당하는 라이선스 상태를 엑셀에 저장
    def save_license_status_to_excel(self, license_status):
        wb = None
        try:
            wb = openpyxl.load_workbook(self.excel_path)
            ws = wb.active

            ip_found = False
            for row in range(2, ws.max_row + 1):
                if ws.cell(row, 1).value == self.current_ip:
                    ws.cell(row, 4, license_status)  # D열(4번째 열)에 저장
                    ip_found = True
                    break

            if not ip_found:
                new_row = ws.max_row + 1
                ws.cell(new_row, 1, self.current_ip)
                ws.cell(new_row, 4, license_status)

            wb.save(self.excel_path)
            return True

        except PermissionError:
            self.util.write_log(
                "라이선스 상태 저장 실패: 엑셀 파일 열림",
                f"ip: {self.current_ip}\nlicense_status: {license_status}\nexcel_path: {self.excel_path}",
            )
            self.show_message("showerror", "저장 오류", "엑셀 파일(esu.xlsx)이 열려있습니다.\n파일을 닫고 다시 시도해주세요.")
            return False
        except Exception as e:
            self.util.write_log(
                "라이선스 상태 저장 중 예외 발생",
                f"ip: {self.current_ip}\nlicense_status: {license_status}\nexcel_path: {self.excel_path}\nerror: {type(e).__name__}: {e}",
            )
            self.show_message("showerror", "오류", f"저장 중 오류 발생: {e}")
            return False
        finally:
            if wb is not None:
                wb.close()

    # 현재 PC IP에 해당하는 확인 값 가져오기

    def get_confirm_value_from_excel(self):
        wb = None
        try:
            wb = openpyxl.load_workbook(self.excel_path, data_only=True)
            ws = wb.active

            val = None
            for row in range(2, ws.max_row + 1):
                if ws.cell(row, 1).value == self.current_ip:
                    val = ws.cell(row, 3).value
                    break

            if val:
                return str(val).strip()
            return None

        except PermissionError:
            self.util.write_log(
                "확인 값 읽기 실패: 엑셀 파일 열림",
                f"ip: {self.current_ip}\nexcel_path: {self.excel_path}",
            )
            self.show_message("showerror", "읽기 오류", "엑셀 파일(esu.xlsx)이 열려있습니다.\n파일을 닫고 다시 시도해주세요.")
            return "OPEN_ERROR"
        except Exception as e:
            self.util.write_log(
                "확인 값 읽기 중 예외 발생",
                f"ip: {self.current_ip}\nexcel_path: {self.excel_path}\nerror: {type(e).__name__}: {e}",
            )
            return None
        finally:
            if wb is not None:
                wb.close()
