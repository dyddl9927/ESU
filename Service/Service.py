from tkinter import messagebox
import subprocess
import openpyxl
from openpyxl import Workbook
import os
from pathlib import Path
import re

class Service :
    def __init__(self, root, ui, util):
        self.root = root
        self.ui = ui
        self.util = util
        # 엑셀파일 저장경로
        self.excel_path = r"c:\1312\esu.xlsx"
        # 인증 cdkey 값 입력
        self.cdkey = "cd키값 입력"
        self.current_ip = self.util.get_local_ip()
        self.current_version = self.util.get_windows_build()

    # 엑셀 파일 및 폴더 확인/생성
    def ensure_excel_exists(self):
        folder = Path(r"c:\1312")
        if not folder.exists():
            folder.mkdir(parents=True, exist_ok=True)
        
        try:
            if not os.path.exists(self.excel_path):
                wb = Workbook()
                ws = wb.active
                ws.title = "ESU인증"
                ws['A1'] = 'IP'
                ws['B1'] = 'DTI'
                ws['C1'] = '확인'
                ws['D1'] = 'END'
                ws['E1'] = 'PRE'  # PRE 열 추가
                wb.save(self.excel_path)
            else:
                # 파일이 있으면 헤더 확인
                wb = openpyxl.load_workbook(self.excel_path)
                ws = wb.active
                if ws['D1'].value != 'END':
                    ws['D1'] = 'END'
                if ws['E1'].value != 'PRE':
                    ws['E1'] = 'PRE'
                wb.save(self.excel_path)
                wb.close()
        except PermissionError:
            messagebox.showerror("파일 오류", "엑셀 파일(esu.xlsx)이 열려있습니다.\n파일을 닫고 프로그램을 다시 실행해주세요.")
    
    # cmd명령어 실행
    def run_command(self, command):
        try:
            result = subprocess.run(
                command,
                shell=True,
                capture_output=True,
                text=True,
                encoding='cp949',
                errors='ignore'
            )
            return result.stdout + result.stderr
        except Exception as e:
            return f"오류: {str(e)}"
    
    # A작업: Windows 버전 확인 후 CD키 설치 후 설치ID(DTI)값 수집후 엑셀에 저장
    def check_installation_id(self):
        # 먼저 Windows 버전 확인
        if not self.util.check_windows_version():
            # 버전이 맞지 않으면 PRE 열에 "실패" 저장
            self.save_pre_status("실패")
            messagebox.showerror(
                "버전 불일치", 
                f"Windows 10 버전이 요구사항과 맞지 않습니다.\n\n"
                f"현재 버전: {self.current_version}\n"
                f"필요 버전: 19045.6456\n\n"
                f"엑셀 PRE 열에 '실패'로 기록되었습니다."
            )
            self.ui.status_label.config(text="버전 불일치 - 설치 실패", fg="red")
            return
        
        # 버전이 맞으면 CD 키 확인
        if not self.cdkey:
            messagebox.showwarning("입력 오류", "CD 키를 입력해주세요!")
            return
        
        self.ui.status_label.config(text="설치ID 수집 중...", fg="blue")
        self.root.update()
        
        # CD 키 설치
        cmd_ipk = f'cscript //Nologo c:\\windows\\system32\\slmgr.vbs /ipk {self.cdkey}'
        self.run_command(cmd_ipk)
        
        # DTI 값 가져오기
        cmd_dti = 'cscript //Nologo c:\\windows\\system32\\slmgr.vbs /dti'
        result_dti = self.run_command(cmd_dti)
        
        dti_value = self.extract_dti_value(result_dti)
        
        if dti_value:
            # 엑셀 저장 시도 (성공/실패 반환 받음)
            if self.save_to_excel(dti_value):
                messagebox.showinfo("성공", f"설치ID가 수집되었습니다.\n\nDTI: {dti_value}\n\n엑셀 파일에 저장되었습니다.")
                self.ui.status_label.config(text="설치 완료", fg="green")
            else:
                self.ui.status_label.config(text="저장 실패 (엑셀 열림)", fg="red")
        else:
            messagebox.showerror("오류", f"설치ID 수집 실패\n\n{result_dti}")
            self.ui.status_label.config(text="설치 실패", fg="red")
    
    # PRE 열에 상태 저장(윈도우 버전 불일치 시 사용)
    def save_pre_status(self, status):
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
            wb.close()
            return True
            
        except PermissionError:
            messagebox.showerror("저장 오류", "엑셀 파일(esu.xlsx)이 열려있습니다.\n파일을 닫고 다시 시도해주세요.")
            return False
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 저장 중 오류 발생: {e}")
            return False
    
    # 설치ID(DTI) 값 추출
    def extract_dti_value(self, output):
        lines = output.strip().split('\n')
        for line in lines:
            if '설치 ID' in line or 'Installation ID' in line:
                parts = line.split(':')
                if len(parts) > 1:
                    return parts[1].strip()
        
        # 결과가 숫자만으로 길게 나오거나 하는 경우 그대로 반환
        clean_output = output.strip()
        if len(clean_output) > 40:
            return clean_output
            
        return clean_output
    
    # 엑셀에 IP와 DTI 저장
    def save_to_excel(self, dti_value):
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
            wb.close()
            return True
            
        except PermissionError:
            messagebox.showerror("저장 오류", "엑셀 파일(esu.xlsx)이 열려있습니다.\n파일을 닫고 다시 시도해주세요.")
            return False
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 저장 중 오류 발생: {e}")
            return False
    
    # B작업: 인증 및 라이선스 상태 확인
    def activate_esu(self):
        self.ui.status_label.config(text="인증 진행 중...", fg="blue")
        self.root.update()
        
        # 엑셀에서 확인 값 가져오기
        confirm_value = self.get_confirm_value_from_excel()
        
        if confirm_value == "OPEN_ERROR":
            self.ui.status_label.config(text="엑셀 파일 열려있음", fg="red")
            return

        if not confirm_value:
            messagebox.showerror("오류", f"엑셀 파일에서 IP({self.current_ip})에 해당하는 확인 값을 찾을 수 없습니다.\n'확인' 항목에 값을 입력해주세요.")
            self.ui.status_label.config(text="인증 실패 (값 없음)", fg="red")
            return
        
        # 인증 실행
        cmd_atp = f'cscript //Nologo c:\\windows\\system32\\slmgr.vbs /atp {confirm_value}'
        result_atp = self.run_command(cmd_atp)
        
        # 라이선스 상태 확인
        self.ui.status_label.config(text="라이선스 상태 확인 중...", fg="blue")
        self.root.update()
        
        cmd_dlv = 'cscript //Nologo c:\\windows\\system32\\slmgr.vbs /dlv'
        result_dlv = self.run_command(cmd_dlv)
        
        # QDFWW 키의 라이선스 상태 추출
        license_status = self.extract_license_status(result_dlv)
        
        if license_status:
            # 엑셀에 라이선스 상태 저장
            if self.save_license_status_to_excel(license_status):
                messagebox.showinfo("성공", 
                    f"인증이 완료되었습니다!\n\n"
                    f"인증 결과:\n{result_atp}\n\n"
                    f"라이선스 상태: {license_status}\n\n"
                    f"엑셀 파일 END 열에 저장되었습니다.")
                self.ui.status_label.config(text="인증 및 상태 확인 완료", fg="green")
            else:
                self.ui.status_label.config(text="저장 실패 (엑셀 열림)", fg="red")
        else:
            messagebox.showwarning("결과", 
                f"인증 결과:\n{result_atp}\n\n"
                f"QDFWW 키의 라이선스 상태를 찾을 수 없습니다.")
            self.ui.status_label.config(text="인증 완료 (상태 미확인)", fg="orange")
    
    # QDFWW 키의 라이선스 상태 추출
    def extract_license_status(self, dlv_output):
        # 한글 윈도우 패턴
        pattern_kor = r"부분 제품 키\s*:\s*QDFWW.*?라이선스 상태\s*:\s*([^\r\n]+)"
        match = re.search(pattern_kor, dlv_output, re.DOTALL | re.IGNORECASE)
        if match:
            return match.group(1).strip()
            
        # 영문 윈도우 패턴
        pattern_eng = r"Partial Product Key\s*:\s*QDFWW.*?License Status\s*:\s*([^\r\n]+)"
        match_eng = re.search(pattern_eng, dlv_output, re.DOTALL | re.IGNORECASE)
        if match_eng:
            return match_eng.group(1).strip()
        
        return None
    
    # 현재 PC IP에 해당하는 라이선스 상태를 엑셀에 저장
    def save_license_status_to_excel(self, license_status):
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
            wb.close()
            return True
            
        except PermissionError:
            messagebox.showerror("저장 오류", "엑셀 파일(esu.xlsx)이 열려있습니다.\n파일을 닫고 다시 시도해주세요.")
            return False
        except Exception as e:
            messagebox.showerror("오류", f"저장 중 오류 발생: {e}")
            return False
    
    # 현재 PC IP에 해당하는 확인 값 가져오기
    def get_confirm_value_from_excel(self):
        try:
            wb = openpyxl.load_workbook(self.excel_path, data_only=True)
            ws = wb.active
            
            val = None
            for row in range(2, ws.max_row + 1):
                if ws.cell(row, 1).value == self.current_ip:
                    val = ws.cell(row, 3).value
                    break
            
            wb.close()
            if val:
                return str(val).strip()
            return None
            
        except PermissionError:
            messagebox.showerror("읽기 오류", "엑셀 파일(esu.xlsx)이 열려있습니다.\n파일을 닫고 다시 시도해주세요.")
            return "OPEN_ERROR"
        except Exception as e:
            return None
