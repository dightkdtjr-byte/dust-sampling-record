#!/bin/bash
# 검증 한 회차: 원본 → 서식 정리 → 기준 시트(완성본) → 프로그램 → 다시 계산·오류 → 9가지 기준 → 원문 대조 → 경계값 → 초과·산소 보정 → 브라우저
#   tools/labform/verify_round.sh <회차 이름>
set -u
cd "$(dirname "$0")"
N=$1; W=work/rounds/$N; S=/tmp/claude-0/-home-user-dust-sampling-record/1c05a7b3-347e-5a3c-9019-6431ab941f16/scratchpad
rm -rf $W; mkdir -p $W
log(){ echo "[$(date +%H:%M:%S)] $*" | tee -a $W/summary.txt; }
log "회차 $N 시작"
python3 -I run_all.py work/all_zip $W/out > $W/run_all.log 2>&1; log "1 서식 정리: $(grep -c '^ok' $W/run_all.log)/19, 경고 $(grep -c '!!' $W/run_all.log)"
cp work/ref_rules.json $W/ 2>/dev/null
python3 -I apply_master.py $W/out $W/golden $W/master.json > $W/apply.log 2>&1; log "2 기준 시트: $(tail -1 $W/apply.log)"
python3 -I check_master.py $W/master.json > $W/check.log 2>&1; log "3 원문 대조: $(grep '판정 합계' $W/check.log)"
python3 -I evidence.py work/ref_zip $W/master_checked.json $W/evidence.json > $W/evidence.log 2>&1
python3 -I build_program.py $W/golden $W/master_checked.json $S/npm/package/dist/jszip.min.js "$W/시험기록부.html" $W/evidence.json > $W/build.log 2>&1; log "4 프로그램: $(tail -1 $W/build.log)"
./lo.sh $W/golden pdf > /dev/null 2>&1
python3 -I errs.py $W/golden/lo > $W/errs.log 2>&1; log "5 오류 칸: $(grep -v '오류 칸 0' $W/errs.log | wc -l)개 파일에 오류 $(grep -v '오류 칸 0' $W/errs.log | tr '\n' ' ' | cut -c1-200)"
python3 -I verify.py $W/golden --pdf $W/golden/pdf > $W/verify.log 2>&1; log "6 9가지 기준: $(sed -n '/^①/,$p' $W/verify.log | grep -E '^[①-⑨]' | tr '\n' ' ')"
python3 -I boundary_test.py $W/golden $W/master.json $W/btest > $W/boundary.log 2>&1; log "7 경계값: $(grep '경계값 시험' $W/boundary.log)"
python3 -I excess_test.py $W/golden $W/master.json $W/xtest > $W/excess.log 2>&1; log "8 초과·산소: $(grep '시험:' $W/excess.log)"
rm -rf $W/e2e; mkdir -p $W/e2e; (cd $S && timeout 300 node e2e.js "$OLDPWD/$W/시험기록부.html" "$OLDPWD/$W/e2e") > $W/e2e.log 2>&1
log "9 브라우저: 오류 $(python3 -c "import json,sys; d=json.load(open('$W/e2e.log')); print(len(d['errors']), '| 시험해보기', d['try1'], d['try2'], '| 내려받기', 'O' if d.get('download') else 'X')" 2>/dev/null || echo '실패')"
log "회차 $N 끝"
