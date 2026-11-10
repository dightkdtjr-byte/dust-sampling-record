# STACKPILOT 로고·아이콘: 5x7 격자 글자를 "블록"으로 쌓아 그린다.
# 모든 모서리는 직각, 색은 단색(그라데이션·둥근 모서리 없음). 칸 하나하나가 위·왼쪽 빛, 아래·오른쪽 그늘을 가진 블록.
import sys
G = {
'S': ["#####","#....","#....","#####","....#","....#","#####"],
'T': ["#####","..#..","..#..","..#..","..#..","..#..","..#.."],
'A': ["..#..",".###.","##.##","#...#","#####","#...#","#...#"],
'C': ["#####","#....","#....","#....","#....","#....","#####"],
'K': ["#...#","#..##","#.##.","###..","#.##.","#..##","#...#"],
'P': ["#####","#...#","#...#","#####","#....","#....","#...."],
'I': ["#####","..#..","..#..","..#..","..#..","..#..","#####"],
'L': ["#....","#....","#....","#....","#....","#....","#####"],
'O': ["#####","#...#","#...#","#...#","#...#","#...#","#####"],
}
def layout(lines, cw, ch, gap, linegap):
    widths = [len(w)*5*cw + (len(w)-1)*gap for w in lines]
    W = max(widths); H = len(lines)*7*ch + (len(lines)-1)*linegap
    cells = []
    for li, word in enumerate(lines):
        x0 = (W - widths[li]) // 2; y0 = li*(7*ch + linegap)
        for letter in word:
            for r, line in enumerate(G[letter]):
                for c, v in enumerate(line):
                    if v == '#': cells.append((x0 + c*cw, y0 + r*ch))
            x0 += 5*cw + gap
    return cells, W, H
def blocks(cells, cw, ch, pal, out, bev, sh):
    R = lambda x, y, w, h: f'<rect x="{x}" y="{y}" width="{w}" height="{h}"/>'
    o = []
    o.append(f'<g fill="{pal["shadow"]}" stroke="{pal["shadow"]}" stroke-width="{out*2}" transform="translate({sh} {sh})">' + ''.join(R(x, y, cw, ch) for x, y in cells) + '</g>')
    o.append(f'<g fill="{pal["line"]}" stroke="{pal["line"]}" stroke-width="{out*2}">' + ''.join(R(x, y, cw, ch) for x, y in cells) + '</g>')
    o.append(f'<g fill="{pal["fill"]}">' + ''.join(R(x, y, cw, ch) for x, y in cells) + '</g>')
    # 칸마다: 위·왼쪽 빛, 아래·오른쪽 그늘 → 블록을 쌓은 모양
    o.append(f'<g fill="{pal["hi"]}">' + ''.join(R(x, y, cw, bev) + R(x, y, bev, ch) for x, y in cells) + '</g>')
    o.append(f'<g fill="{pal["lo"]}">' + ''.join(R(x, y + ch - bev, cw, bev) + R(x + cw - bev, y, bev, ch) for x, y in cells) + '</g>')
    return ''.join(o)
def wordmark(pal, cw=18, ch=16, gap=20, pad=10, out=4, bev=3, sh=5):
    cells, W, H = layout(['STACKPILOT'], cw, ch, gap, 0)
    TW, TH = W + pad*2 + sh, H + pad*2 + sh
    return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {TW} {TH}" width="{TW}" height="{TH}" shape-rendering="crispEdges"><title>STACKPILOT</title>'
            f'<g transform="translate({pad} {pad})">{blocks(cells, cw, ch, pal, out, bev, sh)}</g></svg>')
def icon(lines, pal, cw, ch, gap, linegap, out, bev, sh, size=512, step=32):
    cells, W, H = layout(lines, cw, ch, gap, linegap)
    ox = (size - W - sh) // 2; oy = (size - H - sh) // 2
    # 바탕: 모서리를 한 칸씩 계단처럼 깎은 네모 (둥글리지 않는다)
    s, e = step, size - step
    tile = f'M{s*2},{s} H{e-s} V{s*2} H{e} V{e-s} H{e-s} V{e} H{s*2} V{e-s} H{s} V{s*2} H{s*2} Z'
    return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {size} {size}" width="{size}" height="{size}" shape-rendering="crispEdges">'
            f'<path d="{tile}" fill="{pal["bg"]}" stroke="{pal["rim"]}" stroke-width="10" stroke-linejoin="miter"/>'
            f'<g transform="translate({ox} {oy})">{blocks(cells, cw, ch, pal, out, bev, sh)}</g></svg>')
PALS = {
 'green': dict(fill='#4fe04a', hi='#a8ff94', lo='#27a23a', line='#0f4a1c', shadow='#0a2f14'),
 'clay':  dict(fill='#c8643c', hi='#e9a383', lo='#9c4826', line='#4a2010', shadow='#2e140a'),
}
d = sys.argv[1]
for name, pal in PALS.items():
    open(f'{d}/mainlogo-{name}.svg', 'w').write(wordmark(pal))
ICON = dict(PALS['green'], bg='#102a1e', rim='#2f7a4c', line='#06200f', shadow='#04140a')
open(f'{d}/icon-full.svg', 'w').write(icon(['STACK', 'PILOT'], ICON, cw=15, ch=15, gap=11, linegap=36, out=4, bev=3, sh=5))
open(f'{d}/icon-small.svg', 'w').write(icon(['SP'], ICON, cw=32, ch=36, gap=34, linegap=0, out=10, bev=8, sh=10))
