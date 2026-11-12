import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import './index.css'
import App from './App.jsx'

createRoot(document.getElementById('root')).render(
  <StrictMode>
    <App />
  </StrictMode>,
)

if ('serviceWorker' in navigator && /^https?:$/.test(window.location.protocol)) {
  window.addEventListener('load', () => {
    const swUrl = `${import.meta.env.BASE_URL}sw.js`
    navigator.serviceWorker.register(swUrl, { scope: import.meta.env.BASE_URL }).catch((error) => {
      console.error('Service worker registration failed:', error)
    })
  })
}

// 브라우저가 저장 공간이 부족할 때 이 사이트 데이터를 먼저 지우지 않도록 "계속 보관"을 요청한다
// (사용자가 직접 방문 기록·웹사이트 데이터를 지우는 것은 막을 수 없다)
if (navigator.storage && typeof navigator.storage.persist === 'function') {
  navigator.storage.persisted().then((already) => (already ? true : navigator.storage.persist())).catch(() => {})
}
