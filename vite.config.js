import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  base: '/excel-splitter/', // 專案名稱
  plugins: [react()],
});
