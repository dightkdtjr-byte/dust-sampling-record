# 물질 이름·시험방법 제목을 대기오염공정시험기준 표기로 맞춘다.
#   · 물질 이름: 원문에서 확인한 표기 (1,3-부타디엔은 총칙 표 1의 "1,3-뷰타다이엔". 시험법 본문은 "1.3-"으로 쓰나 이름 표는 총칙.
#     아세트알데히드는 원문 표에 영문만 있어 같은 원문의 "폼알데하이드" 표기법을 따른다)
#   · 글자 칸·수식 안 글자·시트 이름·인쇄 범위 이름·차트 참조·문서 정보까지 한꺼번에 바꾼다
#     (시트 이름과 요약표 이름이 따로 놀면 INDIRECT·매크로가 시트를 못 찾는다)
#   · 시험방법 칸: "ES 번호 + 제목"의 제목을 받은 원문 파일의 제목으로
import re, os, glob

NAME_MAP = [
    ('1,3-부타디엔', '1,3-뷰타다이엔'),     # 총칙 1.2.1 표 1 (별표8 이름 → 공정시험기준 이름)
    ('1.3-뷰타다이엔', '1,3-뷰타다이엔'),
    ('1,2-디클로로에탄', '1,2-다이클로로에테인'),
    ('디클로로메탄', '다이클로로메테인'),
    ('트리클로로에틸렌', '트라이클로로에틸렌'),
    ('클로로포름', '클로로폼'),
    ('아크릴로니트릴', '아크릴로나이트릴'),
    ('스틸렌', '스타이렌'),
    ('염화비닐', '염화바이닐'),
    ('불소화합물', '플루오린화합물'),
    ('브롬화합물', '브로민화합물'),
    ('시안화수소', '사이안화수소'),
    ('포름알데히드', '폼알데하이드'),
    ('포름알데하이드', '폼알데하이드'),
    ('아세트알데히드', '아세트알데하이드'),
    ('크롬화합물', '크로뮴화합물'),
    ('유도결합플라즈마', '유도결합플라스마'),
]
FILE_MAP = {'불소화합물': '플루오린화합물', '브롬': '브로민화합물', '시안화수소': '사이안화수소', '포름알데하이드': '폼알데하이드'}


# 파일별로 고칠 시트 이름: 이황화메틸 기록부의 시트 이름이 원본에서 "이황화탄소"로 복사된 채였다
SHEET_FIX = {'이황화메틸': [('시험기록부-이황화탄소', '시험기록부-이황화메틸')]}


def rename_text(s, pairs=None):
    for a, b in (NAME_MAP if pairs is None else pairs):
        s = s.replace(a, b)
    return s


def pairs_for(path):
    base = os.path.basename(path)
    return NAME_MAP + [p for k, v in SHEET_FIX.items() if k in base for p in v]


def rename_book(b, pairs=None):
    """통합 문서 전체(글자·수식·시트 이름·이름 정의·차트·문서 정보)에서 물질 이름을 바꾼다. 바꾼 개수"""
    NAME_MAP_ = NAME_MAP if pairs is None else pairs
    rename_text_ = lambda s: rename_text(s, NAME_MAP_)
    n = 0
    for p in list(b.parts):
        if not (p.endswith('.xml') or p.endswith('.rels')) or p.startswith('xl/media/'):
            continue
        try:
            x = b.text(p)
        except UnicodeDecodeError:
            continue
        y = rename_text_(x)
        if y != x:
            n += sum(x.count(a) for a, _ in NAME_MAP_)
            b.put(p, y)
    # Book이 들고 있는 시트 이름 목록·openpyxl 사본도 새 이름으로
    b.sheets = {rename_text_(k): v for k, v in b.sheets.items()}
    for ws in b.wb.worksheets:
        nt = rename_text_(ws.title)
        if nt != ws.title:
            ws.title = nt
    return n


def official_titles(ref_dir):
    """ES 번호 → 원문 제목. 파일 이름에는 "/"를 쓸 수 없어 빠지므로(자외선/가시선분광법 → 자외선가시선분광법)
    PDF 첫 쪽에서 파일 이름 제목과 글자(빈칸·/·-·– 빼고)가 같은 줄을 찾아 그 표기를 쓴다"""
    import subprocess
    out = {}
    norm = lambda t: re.sub(r'[\s/\-–―−·]', '', t)
    for p in glob.glob(os.path.join(ref_dir, 'ES *.pdf')):
        n = os.path.splitext(os.path.basename(p))[0]
        m = re.match(r'ES (\d{5}(?:\.\w+)?)[ _](.+)$', n)
        if not m:
            continue
        title = m.group(2).replace('_', '/').strip()
        try:
            lines = [l.strip() for l in subprocess.run(['pdftotext', '-layout', '-f', '1', '-l', '1', p, '-'], capture_output=True, text=True, timeout=30).stdout.splitlines() if l.strip()]
        except Exception:
            lines = []
        # 첫 쪽 제목: "대기오염공정시험기준 ES …" 줄과 연도 줄을 빼고, 영문 "(" 줄 앞까지의 한글 줄을 잇는다
        body = []
        for l in lines:
            if l.startswith('('):
                break
            if '대기오염공정시험기준' in l or re.fullmatch(r'\d{4}', l):
                continue
            body.append(re.sub(r'\s{2,}\d{4}$', '', l))      # 같은 줄 오른쪽 끝의 연도
        seg = re.sub(r'\s*[–―−]\s*', ' - ', re.sub(r'\s+', ' ', ' '.join(body))).strip()
        found = seg if seg and norm(seg) == norm(title) else None
        if not found:
            found = title.replace('플라스마원자발광', '플라스마/원자발광')
        out[m.group(1)] = found
    return out


def fix_method(text, titles):
    """시험방법 글: ES 번호 띄어쓰기를 맞추고, 번호 뒤 제목을 원문 제목으로 (제목이 없으면 붙인다)"""
    def one(m):
        code = m.group(1)
        t = titles.get(code)
        return f'ES {code} {t}' if t else f'ES {code}' + (m.group(2) or '')
    # "ES 01510.1d 배출가스 중 …" 에서 제목은 다음 "(" 짝이 닫히는 ")"나 줄바꿈 앞까지
    return re.sub(r'ES\s*(\d{5}\.\w+)((?:\s+[^\n()]*?(?:\([^\n()]*\)[^\n()]*?)*)?)(?=\)|\n|$)', one, text)
