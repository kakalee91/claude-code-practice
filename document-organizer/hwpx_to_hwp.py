"""
HWPX -> HWP 일괄 변환 스크립트 (hwp_to_hwpx.py의 반대 방향)

사전 준비 (본인 Windows PC에서, 한글/한컴오피스가 설치되어 있어야 합니다):
    pip install pywin32

사용법:
    python hwpx_to_hwp.py "C:\\변환할파일들이있는폴더"
    python hwpx_to_hwp.py "C:\\변환할파일들이있는폴더" --recursive   (하위 폴더까지 포함)
    python hwpx_to_hwp.py "폴더1" "폴더2" "폴더3"                  (여러 폴더 한 번에)

주의:
    - 원본 .hwpx 파일은 그대로 두고, 같은 이름의 .hwp 파일을 새로 만듭니다.
    - 이미 .hwp로 변환된 파일은 건너뜁니다(다시 실행해도 안전).
    - 실행 중 한글 프로그램이 보안 확인 팝업을 띄우며 멈추는 경우가 있습니다.
      그럴 때는 "한글 자동화 보안모듈"로 검색해서 나오는 레지스트리 등록 방법을
      적용하거나, 처음 몇 개 파일은 팝업이 뜰 때 수동으로 눌러 넘겨주세요.
"""

import sys
import os
import glob
import argparse

try:
    import win32com.client as win32
    import pythoncom
except ImportError:
    sys.exit("pywin32가 설치되어 있지 않습니다. 먼저 'pip install pywin32'를 실행하세요.")


def collect_todo(path, recursive=False):
    """path는 폴더일 수도, .hwpx 파일 하나일 수도 있다 (드래그 앤 드롭 대응)."""
    if os.path.isfile(path):
        files = [path] if path.lower().endswith(".hwpx") else []
    else:
        pattern = "**/*.hwpx" if recursive else "*.hwpx"
        files = glob.glob(os.path.join(path, pattern), recursive=recursive)

    todo = []
    for f in files:
        out_path = os.path.splitext(f)[0] + ".hwp"
        if os.path.exists(out_path):
            print(f"[건너뜀] 이미 존재함: {out_path}")
        else:
            todo.append(f)
    return todo


def convert_files(todo):
    if not todo:
        print("변환할 .hwpx 파일을 찾지 못했습니다.")
        return

    print(f"총 {len(todo)}개 파일 변환을 시작합니다.\n")

    pythoncom.CoInitialize()
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")

    # 보안 확인 팝업 우회 시도 (실패해도 무시하고 진행)
    try:
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
    except Exception:
        pass

    try:
        hwp.XHwpWindows.Item(0).Visible = False
    except Exception:
        pass

    success, failed = [], []

    for path in todo:
        out_path = os.path.splitext(path)[0] + ".hwp"
        try:
            hwp.Open(path)
            hwp.SaveAs(out_path, "HWP")
            hwp.Clear(1)
            success.append(path)
            print(f"[완료] {path}")
        except Exception as e:
            failed.append((path, str(e)))
            print(f"[실패] {path} -> {e}")

    hwp.Quit()

    print(f"\n총 {len(todo)}개 중 성공 {len(success)}개, 실패 {len(failed)}개")
    if failed:
        print("\n실패한 파일 목록:")
        for path, err in failed:
            print(f" - {path}: {err}")


def run(paths, recursive=False):
    """GUI 등 다른 스크립트에서 그대로 불러 쓰기 위한 진입점.
    paths: 폴더 경로 또는 .hwpx 파일 경로의 리스트 (섞여 있어도 됨)."""
    all_todo = []
    for path in paths:
        all_todo.extend(collect_todo(path, recursive))
    convert_files(all_todo)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="HWPX 파일을 HWP로 일괄 변환")
    parser.add_argument("folders", nargs="+", help="변환할 .hwpx 파일이 들어있는 폴더 경로 (여러 개 가능)")
    parser.add_argument("--recursive", action="store_true", help="하위 폴더까지 포함")
    args = parser.parse_args()

    run(args.folders, args.recursive)
