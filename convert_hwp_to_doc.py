import subprocess
import os
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

def convert_hwp_to_doc(hwp_file_path, output_dir):
    """
    LibreOffice CLI를 사용하여 HWP 파일을 DOC로 변환한다.
    """
    # LibreOffice 설치 경로 찾기 (Windows)
    libreoffice_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]

    soffice = None
    for path in libreoffice_paths:
        if os.path.exists(path):
            soffice = path
            break

    if not soffice:
        print("❌ LibreOffice가 설치되어 있지 않습니다.")
        print("LibreOffice를 설치하고 다시 시도하세요.")
        return False

    # 파일 존재 확인
    if not os.path.exists(hwp_file_path):
        print(f"❌ 파일을 찾을 수 없습니다: {hwp_file_path}")
        return False

    # 출력 디렉토리 생성
    os.makedirs(output_dir, exist_ok=True)

    # LibreOffice 변환 명령 실행
    cmd = [
        soffice,
        "--headless",
        "--convert-to", "doc",
        "--outdir", output_dir,
        hwp_file_path
    ]

    print(f"🔄 변환 진행 중: {Path(hwp_file_path).name} → DOC")
    print(f"명령: {' '.join(cmd)}")

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)

        if result.returncode == 0:
            output_filename = Path(hwp_file_path).stem + ".doc"
            output_path = os.path.join(output_dir, output_filename)

            if os.path.exists(output_path):
                print(f"✔ 변환 완료!")
                print(f"출력 파일: {output_path}")
                return True
            else:
                print(f"❌ 변환 후 파일이 생성되지 않았습니다.")
                return False
        else:
            print(f"❌ 변환 실패 (반환 코드: {result.returncode})")
            print(f"오류: {result.stderr}")
            return False

    except subprocess.TimeoutExpired:
        print("❌ 변환 시간 초과 (60초)")
        return False
    except Exception as e:
        print(f"❌ 변환 중 오류 발생: {e}")
        return False

if __name__ == '__main__':
    hwp_file = r"d:\Dev\Project\sample\정보통신설비_2026성능점검결과서_엠코지니어스타.hwp"
    output_dir = r"d:\Dev\Project\sample\doc"

    success = convert_hwp_to_doc(hwp_file, output_dir)
    sys.exit(0 if success else 1)
