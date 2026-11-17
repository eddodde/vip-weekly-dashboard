# -*- coding: utf-8 -*-
"""convert.ps1이 만든 CSV들을 시드(data/perf_long.csv)에 병합한다.

사용:
    python tools/merge_seed.py <변환 CSV 폴더>            # 점검만(dry run)
    python tools/merge_seed.py <변환 CSV 폴더> --write    # 시드에 기록(직전 시드는 .bak_날짜로 백업)

규칙:
  - 새 파일에 든 (grain, perspective, year) 블록은 시드에서 통째로 교체한다.
  - 그 밖의 블록은 그대로 둔다. 마지막에 키 중복을 제거한다(새 값 우선).
"""
import csv
import glob
import os
import shutil
import sys
import datetime
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SEED = os.path.join(ROOT, "data", "perf_long.csv")
KEY = ["grain", "perspective", "year", "period", "seg1", "seg2", "metric"]


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not args:
        print(__doc__)
        return 1
    files = sorted(glob.glob(os.path.join(args[0], "*.csv")))
    if not files:
        print("변환 CSV가 없습니다:", args[0])
        return 1
    seed = pd.read_csv(SEED, dtype=str)
    new = pd.concat([pd.read_csv(f, dtype=str) for f in files], ignore_index=True)
    if list(seed.columns) != list(new.columns):
        print("열 구성이 시드와 다릅니다:", list(new.columns))
        return 1
    nk = set(map(tuple, new[["grain", "perspective", "year"]].drop_duplicates().values.tolist()))
    keep = seed[[k not in nk for k in map(tuple, seed[["grain", "perspective", "year"]].values.tolist())]]
    merged = pd.concat([keep, new], ignore_index=True)
    merged = merged.drop_duplicates(subset=KEY, keep="last").reset_index(drop=True)

    print("파일          :", [os.path.basename(f) for f in files])
    print("교체 블록     :", sorted(nk))
    print("시드 %d행 → 병합 %d행 (새 %d행)" % (len(seed), len(merged), len(new)))
    day = merged[(merged.grain == "day") & (merged.perspective == "overall") & (merged.metric == "DAU")]
    for yr in sorted(day.year.unique())[-1:]:
        d = day[day.year == yr].sort_values("period_sort")
        print("일자 실적 마지막 5일(%s): %s" % (yr, ", ".join(
            "%s %s" % (p, format(round(float(v)), ",")) for p, v in zip(d.period.tail(5), d.value.tail(5)))))

    if "--write" in sys.argv:
        bak = SEED + ".bak_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy2(SEED, bak)
        merged.to_csv(SEED, index=False, quoting=csv.QUOTE_ALL, encoding="utf-8-sig")
        print("기록 완료:", SEED, "(백업:", os.path.basename(bak) + ")")
    else:
        print("(점검만 했습니다 — 기록하려면 --write)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
