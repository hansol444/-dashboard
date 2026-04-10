"""
Placement Survey 전체 파이프라인 자동 실행 에이전트

사용법:
  python run_placement_agent.py --quarter 26Q1
  python run_placement_agent.py --quarter 26Q1 --dry-run   # 실행 없이 경로 확인만

파이프라인:
  Stage 1: run_jk.py + run_am.py  → R_통합 생성
  Stage 2: calc_rms.py + calc_rms_am.py  → RMS 계산
  Stage 3: gen_ppt.py  → PPT 생성 (파일 경로/분기 자동 패치)
"""

import sys
import re
import argparse
import subprocess
import shutil
from pathlib import Path

BASE_DIR = Path(r"C:\Users\olivia408\projects\-dashboard\Placement survey")

# 전체 분기 목록 (필요 시 여기에 추가)
ALL_JK_QUARTERS = [
    '23Q1','23Q2','23Q3','23Q4',
    '24Q1','24Q2','24Q3','24Q4',
    '25Q1','25Q2','25Q3','25Q4',
    '26Q1','26Q2','26Q3','26Q4',
]
ALL_AM_QUARTERS = [
    '23.1Q','23.2Q','23.3Q','23.4Q',
    '24.1Q','24.2Q','24.3Q','24.4Q',
    '25.1Q','25.2Q','25.3Q','25.4Q',
    '26.1Q','26.2Q','26.3Q','26.4Q',
]


def jk_to_am_quarter(jk_q: str) -> str:
    """'26Q1' → '26.1Q'"""
    m = re.match(r'(\d{2})Q(\d)', jk_q)
    if not m:
        raise ValueError(f"분기 형식 오류: {jk_q} (예: 26Q1)")
    return f"{m.group(1)}.{m.group(2)}Q"


def prev_jk_quarter(jk_q: str) -> str:
    if jk_q not in ALL_JK_QUARTERS:
        raise ValueError(f"{jk_q} 가 분기 목록에 없습니다. ALL_JK_QUARTERS를 업데이트하세요.")
    idx = ALL_JK_QUARTERS.index(jk_q)
    if idx == 0:
        raise ValueError("이전 분기가 없습니다.")
    return ALL_JK_QUARTERS[idx - 1]


def rolling_12(all_quarters: list, current: str) -> list:
    """현재 분기 포함 최근 12개 반환"""
    if current not in all_quarters:
        raise ValueError(f"{current} 가 분기 목록에 없습니다.")
    idx = all_quarters.index(current)
    start = max(0, idx - 11)
    return all_quarters[start:idx + 1]


def run_step(name: str, cmd: str, cwd: Path, dry_run: bool):
    print(f"\n{'='*60}")
    print(f"  STEP: {name}")
    print(f"  CMD : {cmd}")
    print(f"  CWD : {cwd}")
    print('='*60)
    if dry_run:
        print("  [dry-run] 스킵")
        return
    result = subprocess.run(cmd, cwd=str(cwd), shell=True)
    if result.returncode != 0:
        print(f"\n[ERROR] {name} 실패 (exit {result.returncode})")
        sys.exit(result.returncode)
    print(f"  ✓ {name} 완료")


def patch_gen_ppt(quarter: str, jk_quarters: list, am_quarters: list):
    """gen_ppt.py 의 파일 경로·분기 목록을 현재 분기에 맞게 패치 (백업 후 교체)"""
    gen_ppt = BASE_DIR / "gen_ppt.py"
    backup = BASE_DIR / "gen_ppt.py.bak"

    # 백업
    shutil.copy2(gen_ppt, backup)

    prev_q = prev_jk_quarter(quarter)
    am_q = jk_to_am_quarter(quarter)

    template_name = f"Placement Survey_{prev_q}.pptx"
    jk_rms_name  = f"{quarter}_JK_RMS_v1.xlsx"
    am_rms_name  = f"{quarter}_AM_RMS_v1.xlsx"
    output_name  = f"Placement Survey_{quarter}_output1.pptx"

    content = gen_ppt.read_text(encoding="utf-8")

    content = re.sub(
        r'^TEMPLATE = .*$',
        f'TEMPLATE = BASE_DIR / "{template_name}"',
        content, flags=re.MULTILINE
    )
    content = re.sub(
        r'^JK_RMS = .*$',
        f'JK_RMS = BASE_DIR / "JK 전체" / "{jk_rms_name}"',
        content, flags=re.MULTILINE
    )
    content = re.sub(
        r'^AM_RMS = .*$',
        f'AM_RMS = BASE_DIR / "AM 전체" / "{am_rms_name}"',
        content, flags=re.MULTILINE
    )
    content = re.sub(
        r'^OUTPUT = .*$',
        f'OUTPUT = BASE_DIR / "{output_name}"',
        content, flags=re.MULTILINE
    )
    content = re.sub(
        r'^JK_QUARTERS = \[.*?\]',
        f"JK_QUARTERS = {jk_quarters}",
        content, flags=re.MULTILINE | re.DOTALL
    )
    content = re.sub(
        r'^AM_QUARTERS = \[.*?\]',
        f"AM_QUARTERS = {am_quarters}",
        content, flags=re.MULTILINE | re.DOTALL
    )

    gen_ppt.write_text(content, encoding="utf-8")
    print(f"  gen_ppt.py 패치 완료 (백업: gen_ppt.py.bak)")
    print(f"    TEMPLATE = {template_name}")
    print(f"    JK_RMS   = JK 전체/{jk_rms_name}")
    print(f"    AM_RMS   = AM 전체/{am_rms_name}")
    print(f"    OUTPUT   = {output_name}")
    print(f"    JK_QUARTERS (최근 12개): {jk_quarters}")

    return output_name


def restore_gen_ppt():
    backup = BASE_DIR / "gen_ppt.py.bak"
    if backup.exists():
        shutil.copy2(backup, BASE_DIR / "gen_ppt.py")
        backup.unlink()
        print("  gen_ppt.py 복원 완료")


def main():
    parser = argparse.ArgumentParser(description="Placement Survey 파이프라인 자동 실행")
    parser.add_argument("--quarter", required=True, help="실행할 분기 (예: 26Q1)")
    parser.add_argument("--dry-run", action="store_true", help="경로·명령 확인만 (실제 실행 없음)")
    parser.add_argument("--stage", type=int, choices=[1, 2, 3], help="특정 Stage만 실행 (기본: 전체)")
    args = parser.parse_args()

    q = args.quarter
    dry = args.dry_run
    am_q = jk_to_am_quarter(q)
    prev_q = prev_jk_quarter(q)
    prev_am_q = jk_to_am_quarter(prev_q)

    jk_quarters = rolling_12(ALL_JK_QUARTERS, q)
    am_quarters = rolling_12(ALL_AM_QUARTERS, am_q)

    print(f"\n{'#'*60}")
    print(f"  Placement Survey Agent  —  분기: {q} / AM: {am_q}")
    print(f"  이전 분기: {prev_q} / AM: {prev_am_q}")
    if dry:
        print("  MODE: dry-run")
    print(f"{'#'*60}")

    # ── 파일 경로 계산 ──
    jk_raw    = BASE_DIR / "JK 전체" / f"{q}_JK_Raw.xlsx"
    jk_result = BASE_DIR / "JK 전체" / f"{q}_JK_결과.xlsx"
    jk_rms    = BASE_DIR / "JK 전체" / f"{q}_JK_RMS_v1.xlsx"

    am_raw    = BASE_DIR / "AM 전체" / f"{q}_AM_Raw.xlsx"
    am_base   = BASE_DIR / "AM 전체" / f"{prev_q}_AM_결과.xlsx"
    am_result = BASE_DIR / "AM 전체" / f"{q}_AM_결과.xlsx"
    am_rms    = BASE_DIR / "AM 전체" / f"{q}_AM_RMS_v1.xlsx"

    # ── 사전 파일 체크 ──
    missing = []
    if not dry:
        if not jk_raw.exists():  missing.append(str(jk_raw))
        if not am_raw.exists():  missing.append(str(am_raw))
        if not am_base.exists(): missing.append(str(am_base))

    if missing:
        print("\n[ERROR] 아래 파일이 없습니다:")
        for f in missing:
            print(f"  - {f}")
        print("\n파일을 준비한 뒤 다시 실행하세요.")
        sys.exit(1)

    run_all = args.stage is None

    # ── Stage 1 ──
    if run_all or args.stage == 1:
        run_step(
            "Stage 1-JK: R_통합 생성",
            f'python run_jk.py --raw "JK 전체/{jk_raw.name}" --quarter {q}',
            BASE_DIR, dry
        )
        run_step(
            "Stage 1-AM: R_통합 생성",
            f'python run_am.py --base "AM 전체/{am_base.name}" --raw "AM 전체/{am_raw.name}" --quarter {q}',
            BASE_DIR, dry
        )

    # ── Stage 2 ──
    if run_all or args.stage == 2:
        run_step(
            "Stage 2-JK: RMS 계산",
            f'python calc_rms.py --input "JK 전체/{jk_result.name}" --output "JK 전체/{jk_rms.name}"',
            BASE_DIR, dry
        )
        run_step(
            "Stage 2-AM: RMS 계산",
            f'python calc_rms_am.py --input "AM 전체/{am_result.name}" --output "AM 전체/{am_rms.name}"',
            BASE_DIR, dry
        )

    # ── Stage 3 ──
    if run_all or args.stage == 3:
        if not dry:
            output_name = patch_gen_ppt(q, jk_quarters, am_quarters)
        else:
            output_name = f"Placement Survey_{q}_output1.pptx"
            print(f"\n  [dry-run] gen_ppt.py 패치 스킵")
            print(f"    예상 OUTPUT = {output_name}")

        run_step(
            "Stage 3: PPT 생성",
            "python gen_ppt.py",
            BASE_DIR, dry
        )

        if not dry:
            restore_gen_ppt()
            output_path = BASE_DIR / output_name
            if output_path.exists():
                print(f"\n  ✅ PPT 생성 완료: {output_path}")
            else:
                print(f"\n  ⚠️  PPT 파일 미생성: {output_path}")

    print(f"\n{'#'*60}")
    print(f"  ✅ Placement Survey {q} 파이프라인 완료!")
    if not dry:
        print(f"  수작업 필요: 2-3P 목차, 7-10P/14-17P Scatter, Appendix")
    print(f"{'#'*60}\n")


if __name__ == "__main__":
    main()
