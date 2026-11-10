# STACKPILOT 로고·아이콘: 원래 로고의 픽셀 글자꼴(세로 획 2칸, 가로 획 1칸, 둥근 곳은 한 칸씩 계단)을 격자에 맞춰 다시 그린다.
# 글자는 한 덩어리 단색 면(칸마다 무늬를 넣지 않는다), 직각 모서리, 짙은 윤곽선, 딱 떨어지는 그림자,
# 글자 윗변에만 가는 빛, 아랫변에만 가는 그늘. 그라데이션·둥근 모서리 없음.
import sys
G = {
'S': [".######.","##....##","##......",".######.","......##","......##","##....##",".######."],
'T': ["########","...##...","...##...","...##...","...##...","...##...","...##...","...##..."],
'A': ["...##...","..####..",".##..##.","##....##","##....##","########","##....##","##....##"],
'C': [".######.","##....##","##......","##......","##......","##......","##....##",".######."],
'K': ["##....##","##...##.","##..##..","#####...","#####...","##..##..","##...##.","##....##"],
'P': ["#######.","##....##","##....##","##....##","#######.","##......","##......","##......"],
'I': ["######","..##..","..##..","..##..","..##..","..##..","..##..","######"],
'L': ["##.....","##.....","##.....","##.....","##.....","##.....","##.....","#######"],
'O': [".######.","##....##","##....##","##....##","##....##","##....##","##....##",".######."],
}
ROWS = 8
def layout(lines, cw, ch, gap, linegap):
    width = lambda w: sum(len(G[l][0]) for l in w)*cw + (len(w)-1)*gap
    widths = [width(w) for w in lines]
    W = max(widths); H = len(lines)*ROWS*ch + (len(lines)-1)*linegap
    cells = []
    for li, word in enumerate(lines):
        x0 = (W - widths[li]) // 2; y0 = li*(ROWS*ch + linegap)
        for letter in word:
            g = G[letter]
            for r, line in enumerate(g):
                for c, v in enumerate(line):
                    if v == '#': cells.append((x0 + c*cw, y0 + r*ch, r > 0 and g[r-1][c] == '#', r < ROWS-1 and g[r+1][c] == '#'))
            x0 += len(g[0])*cw + gap
    return cells, W, H
def letters(cells, cw, ch, pal, out, bev, sh):
    R = lambda x, y, w, h: f'<rect x="{x}" y="{y}" width="{w}" height="{h}"/>'
    o = []
    o.append(f'<g fill="{pal["shadow"]}" stroke="{pal["shadow"]}" stroke-width="{out*2}" transform="translate({sh} {sh})">' + ''.join(R(x, y, cw, ch) for x, y, *_ in cells) + '</g>')
    o.append(f'<g fill="{pal["line"]}" stroke="{pal["line"]}" stroke-width="{out*2}">' + ''.join(R(x, y, cw, ch) for x, y, *_ in cells) + '</g>')
    # 면: 칸 사이에 틈이 보이지 않게 같은 색 테두리로 살짝 겹친다
    o.append(f'<g fill="{pal["fill"]}" stroke="{pal["fill"]}" stroke-width="1">' + ''.join(R(x, y, cw, ch) for x, y, *_ in cells) + '</g>')
    o.append(f'<g fill="{pal["hi"]}">' + ''.join(R(x, y, cw, bev) for x, y, up, dn in cells if not up) + '</g>')
    o.append(f'<g fill="{pal["lo"]}">' + ''.join(R(x, y + ch - bev, cw, bev) for x, y, up, dn in cells if not dn) + '</g>')
    return ''.join(o)
def wordmark(pal, cw=12, ch=15, gap=22, pad=10, out=4, bev=4, sh=6):
    cells, W, H = layout(['STACKPILOT'], cw, ch, gap, 0)
    TW, TH = W + pad*2 + sh, H + pad*2 + sh
    return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {TW} {TH}" width="{TW}" height="{TH}" shape-rendering="crispEdges"><title>STACKPILOT</title>'
            f'<g transform="translate({pad} {pad})">{letters(cells, cw, ch, pal, out, bev, sh)}</g></svg>')
def icon(lines, pal, cw, ch, gap, linegap, out, bev, sh, size=512, m=24):
    cells, W, H = layout(lines, cw, ch, gap, linegap)
    ox = (size - W - sh) // 2; oy = (size - H - sh) // 2
    return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {size} {size}" width="{size}" height="{size}" shape-rendering="crispEdges">'
            f'<rect x="{m}" y="{m}" width="{size-2*m}" height="{size-2*m}" fill="{pal["bg"]}"/>'
            f'<rect x="{m+7}" y="{m+7}" width="{size-2*m-14}" height="{size-2*m-14}" fill="none" stroke="{pal["rim"]}" stroke-width="14"/>'
            f'<g transform="translate({ox} {oy})">{letters(cells, cw, ch, pal, out, bev, sh)}</g></svg>')
PALS = {
 'green': dict(fill='#3fd447', hi='#9df58f', lo='#249c38', line='#0c3d18', shadow='#0c3d18'),
 'clay':  dict(fill='#c8643c', hi='#e6a080', lo='#9c4826', line='#3f1b0d', shadow='#3f1b0d'),
}
d = sys.argv[1]
for name, pal in PALS.items():
    open(f'{d}/mainlogo-{name}.svg', 'w').write(wordmark(pal))
ICON = dict(PALS['green'], bg='#ffffff', rim='#3fd447')  # 아이콘 바탕은 흰색
open(f'{d}/icon-full.svg', 'w').write(icon(['STACK', 'PILOT'], ICON, cw=9, ch=13, gap=12, linegap=44, out=4, bev=4, sh=6))
open(f'{d}/icon-small.svg', 'w').write(icon(['SP'], ICON, cw=20, ch=30, gap=36, linegap=0, out=10, bev=8, sh=12, m=16))
