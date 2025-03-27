import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// https://vitejs.dev/config/
export default defineConfig({
  base: './', // 必须包含此语句，否则将导致部署失败
  plugins: [react()],
  server: {
    host: '0.0.0.0',
  }
}); 