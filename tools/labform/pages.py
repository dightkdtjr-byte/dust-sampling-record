# pages.py <pdf폴더> <out.png> [dpi] : 각 pdf의 첫 쪽을 나란히
import sys, glob, os, subprocess
from PIL import Image
d, out = sys.argv[1], sys.argv[2]; dpi = sys.argv[3] if len(sys.argv) > 3 else '50'
pages = sys.argv[4].split(',') if len(sys.argv) > 4 else None
ims = []
for i, f in enumerate(sorted(glob.glob(d + '/*.pdf'))):
    pg = pages[i] if pages else '1'
    base = f[:-4] + '-pg'
    subprocess.run(['pdftoppm', '-r', dpi, '-f', pg, '-l', pg, '-png', '-singlefile', f, base], check=True)
    ims.append(Image.open(base + '.png'))
w = sum(i.size[0] for i in ims) + 10 * len(ims); h = max(i.size[1] for i in ims)
s = Image.new('RGB', (w, h), 'gray'); x = 0
for i in ims: s.paste(i, (x, 0)); x += i.size[0] + 10
s.save(out); print(s.size)
