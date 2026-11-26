#!/bin/bash
# 최종 회차: 최종 코드로 처음부터 + 방식 3(무작위)·5(정적 감사·다른 엔진)·6(인쇄)·7(프로그램↔엑셀)·9(극단값)·10(프로그램 계산)
cd "$(dirname "$0")"
W=work/rounds/final
./verify_round.sh final > /dev/null 2>&1
log(){ echo "[$(date +%H:%M:%S)] $*" | tee -a $W/summary.txt; }
python3 -I random_test.py $W/golden $W/master.json $W/rand 8 99 > $W/random.log 2>&1; log "방식3 무작위: $(head -1 $W/random.log)"
python3 -I audit_round.py $W/golden $W/master.json $W/audit > $W/audit.log 2>&1; log "방식5 정적 감사: $(head -1 $W/audit.log)"
(cd /tmp/claude-0/-home-user-dust-sampling-record/1c05a7b3-347e-5a3c-9019-6431ab941f16/scratchpad/ffp && node engine_eval.js $OLDPWD/$W/audit 2>&1 | grep -v "^ *at \|^#ERROR") > $W/engine.log; log "방식5 다른 엔진: $(head -1 $W/engine.log)"
python3 -I visual_round.py $W/golden $W/master.json $W/visual > $W/visual.log 2>&1; log "방식6 인쇄: $(tail -1 $W/visual.log)"
mkdir -p $W/p7 && (cd /tmp/claude-0/-home-user-dust-sampling-record/1c05a7b3-347e-5a3c-9019-6431ab941f16/scratchpad && node e2e2.js "$OLDPWD/$W/시험기록부.html" "$OLDPWD/$W/p7") > $W/p7.log 2>&1
python3 -I program_vs_excel.py $W/p7 > $W/p7compare.log 2>&1; log "방식7 프로그램↔엑셀: $(head -1 $W/p7compare.log)"
python3 -I edge_test.py $W/golden $W/master.json $W/edge > $W/edge.log 2>&1; log "방식9 극단값: $(head -1 $W/edge.log)"
python3 -I prop_test.py "$W/시험기록부.html" $W/prop > $W/prop.log 2>&1; log "방식10 프로그램 계산: $(head -1 $W/prop.log)"
python3 -I fix_test.py $W/golden $W/fixt > $W/fixt.log 2>&1; log "방식11 원본 수식 고친 것: $(tail -1 $W/fixt.log)"
log "최종 회차 끝"
