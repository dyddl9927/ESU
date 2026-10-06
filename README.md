# ESU

---

## ESU 인증툴

CD 키 설치 및 인증 전에 설정된 ESU 정품 인증 ID의 라이선스 상태를 확인합니다.
이미 인증된 PC는 Windows 빌드 검사, 키 설치, 설치 ID 수집 및 재인증을 생략하고
엑셀 END 열에 `사용 허가됨`을 기록합니다. 키만 설치되고 인증되지 않은 PC는
기존 절차를 진행합니다. 상태 조회 실패 시에는 작업을 중단하고 로그를 기록합니다.
이 검사는 현재 설정된 ESU 연차의 인증 ID를 대상으로 합니다.

설치 허용 빌드는 `19045.6456`, `19045.6466`으로 유지합니다.
키 설치 후 ESU 제품의 부분 제품 키를 확인하고, 설치 ID는 54/63자리 숫자만 저장합니다.
인증 후 상태는 표시 언어와 관계없이 숫자 라이선스 상태로 확인합니다.
엑셀 준비 실패 시 작업을 중단하며, 명령에는 실행 시간 제한을 적용합니다.
설치 및 인증은 백그라운드에서 실행하고 작업 중 인증 버튼을 비활성화합니다.

검증: `python -m unittest discover -s tests -v` (실제 제품 키 설치 없이 모의 검증).

## 구성

-Service

    ensure_excel_exists

        엑셀 파일 및 폴더 확인/생성

    run_command

        cmd명령어 실행

    check_installation_id

        A작업 : Windows 버전 확인후 CD키 설치, 설치ID(DTI)값 수집후 엑셀에 저장

    save_pre_status

        엑셀 PRE열에 상태 저장(윈도우 버전 불일치시 사용)

    extract_dti_value

        설치ID(DTI) 값 추출

    save_to_excel

        엑셀에 IP 와 설치ID(DTI)값 저장
        
    activate_esu

        B작업 : esu인증 및 라이선스 상태 확인

    extract_license_status

        QDFWW 키의 라이선스 상태추출(윈도우의 부분제품키와 라이선스 상태를 가져옴)

    save_license_status_to_excel

        현재 PC IP에 해당하는 라이선스 상태를 엑셀에 저장

    get_confirm_value_from_excel

        현재 PC IP에 해당하는 확인 값 가져오기

-Util

    get_windows_build

        윈도우 빌드 번호 가져오기(레지스트리 값)

    check_windows_version

        윈도우 10버전이 19045.6456인지 확인

    get_lacal_ip

        현재 사용자 IP값 가져오시

    _default_icon_path

        기본 아이콘 경로 반환

    is_admin

        관리자 권한 확인

    request_admin

        관리자 권한으로 재시작

-Ui

    setup_ui

        UI 설정

            메인 프레임

            타이틀

            IP 표시

            Windows 버전 표시

            B버튼 프레임

            상태 표시

    center_window

        창을 화면 중앙에 배치
        
-esu_core

-main
