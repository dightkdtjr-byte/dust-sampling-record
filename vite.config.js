import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import { createHash } from 'node:crypto'
import { readFileSync, readdirSync, writeFileSync, existsSync } from 'node:fs'
import { join } from 'node:path'

// 아이콘 캐시 문제: 브라우저와 휴대폰은 아이콘을 주소 기준으로 오래 기억한다.
// 아이콘 파일 내용으로 만든 값을 주소 뒤에 자동으로 붙여서(?v=…), 아이콘 그림을 바꾸면
// 파일 이름을 손대지 않아도 새 주소가 되어 새 아이콘을 받아 가게 한다.
function iconVersion() {
  const iconFile = (name) => /\.(png|ico)$/.test(name) && /(favicon|apple-touch-icon|android-chrome)/.test(name)
  const version = () => {
    const hash = createHash('sha1')
    readdirSync('public').filter(iconFile).sort().forEach((name) => hash.update(name).update(readFileSync(join('public', name))))
    return hash.digest('hex').slice(0, 10)
  }
  let outDir = 'dist'
  return {
    name: 'icon-version',
    configResolved(config) { outDir = config.build.outDir },
    transformIndexHtml(html) { return html.replace(/\?v=2\b/g, `?v=${version()}`) },
    writeBundle() {
      const v = version()
      for (const name of ['manifest.webmanifest', 'site.webmanifest']) {
        const file = join(outDir, name)
        if (!existsSync(file)) continue
        const manifest = JSON.parse(readFileSync(file, 'utf8'))
        ;(manifest.icons || []).forEach((icon) => { icon.src = `${String(icon.src).split('?')[0]}?v=${v}` })
        writeFileSync(file, JSON.stringify(manifest, null, 2))
      }
    },
  }
}

export default defineConfig({
  base: './',
  plugins: [react(), tailwindcss(), iconVersion()],
  server: {
    host: '0.0.0.0',
    port: 4173,
    allowedHosts: ['.loca.lt', '.trycloudflare.com'],
  },
})
