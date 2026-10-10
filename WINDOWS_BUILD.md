# Windows 실행파일 빌드

대상은 Windows x64이며 결과물은 `ReadQRToxlsx.exe` 한 개입니다. EXE 실행에는 Python 설치가 필요하지 않습니다. 사용자가 앱에서 기존 `.xlsx`를 선택합니다.

## Windows PC에서 빌드

Python 3.12 x64를 설치할 때 Tcl/Tk와 Python Launcher를 포함한 뒤 이 폴더의 `build_windows.bat`을 실행합니다. 배치 파일은 별도 `.venv-build`에 빌드 의존성을 설치하고, PyInstaller로 `dist/ReadQRToxlsx.exe`를 생성하고, 해당 실행파일을 직접 실행하여 검증합니다.

명령으로 실행하려면:

```powershell
py -3.12 -m venv .venv-build
.venv-build\Scripts\python -m pip install -r requirements-build.txt
.venv-build\Scripts\python scripts\build_windows.py
.venv-build\Scripts\python scripts\verify_windows.py
```

## GitHub Actions에서 빌드

`.github/workflows/windows-exe.yml`은 `codex/windows-exe-*` 브랜치로 push되면 Windows 2022 서버에서 자동 실행됩니다. `workflow_dispatch`도 정의되어 있지만 GitHub의 수동 실행 버튼은 워크플로가 기본 브랜치에 있어야 제공될 수 있습니다.

성공한 실행의 **ReadQRToxlsx-Windows-x64** 아티팩트를 내려받아 압축을 풀면 EXE와 `smoke-report.json`이 나옵니다. 아티팩트는 14일 동안 보관합니다.

자동 검사는 생성한 테스트 데이터만 사용합니다:

- Tk 화면 초기화
- Pillow와 zxing-cpp 네이티브 모듈을 사용한 QR 생성·해독
- 2,300행 이후 QR 조회와 5개 항목의 작정금액 확인
- XLSX 입금 기록·백업·저장·재조회, 다른 ZIP 구성 요소 보존
- EXE의 Windows PE 및 x64 형식 확인

실제 스캐너 연결과 사용자 PC에서의 화면·Excel 동작은 배포 후 확인해야 합니다. 코드 서명은 포함하지 않습니다.

## Mac 환경

[PyInstaller는 크로스 컴파일러가 아니므로](https://www.pyinstaller.org/en/stable/) Windows EXE는 Windows에서 빌드합니다. Mac에서 `scripts/build_windows.py`를 실행하면 이를 안내하고 종료합니다.

엑셀·JPEG·백업 등 사용자 자료는 빌드 소스나 EXE에 포함하지 않습니다. 공개 저장소로 빌드할 때에는 앱 모듈·아이콘·의존성·패키징 스크립트·워크플로만 명시적으로 선택해서 업로드해야 합니다. 로컬 참고 자료의 QR ID·금액을 기록한 문서와 통합 검사 파일은 업로드하지 않습니다.
