# LibreOffice 다시 계산을 폴더 여러 개에 동시에 (lo.sh는 폴더마다 따로 임시 프로필을 쓰므로 겹치지 않는다)
import os, subprocess
from concurrent.futures import ThreadPoolExecutor

LO = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'lo.sh')


def run(dirs, pdf=False, jobs=None):
    jobs = jobs or int(os.environ.get('LABFORM_JOBS', min(4, os.cpu_count() or 1)))
    args = lambda d: [LO, d] + (['pdf'] if pdf else [])
    with ThreadPoolExecutor(max(1, jobs)) as ex:
        list(ex.map(lambda d: subprocess.run(args(d), capture_output=True, timeout=1800), dirs))
