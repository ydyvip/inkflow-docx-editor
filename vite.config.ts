import { defineConfig } from 'vite';
import solid from 'vite-plugin-solid';
import tailwindcss from '@tailwindcss/vite';

// https://vite.dev/config/
// base: './' —— 让产物使用相对路径（./assets/...）。
// 浏览器里用 http 服务（npm run dev / preview）时相对路径照常解析；
// Electron 生产环境用 loadFile('file://...') 加载 dist/index.html 时，
// 绝对路径 /assets/... 会错指到磁盘根目录导致 ERR_FILE_NOT_FOUND、应用无法挂载，
// 相对路径则能正确指向 dist/assets/...。
export default defineConfig({
  base: './',
  plugins: [tailwindcss(), solid()],
});
