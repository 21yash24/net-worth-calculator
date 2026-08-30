// @ts-check
import { defineConfig } from 'astro/config';
import tailwindcss from '@tailwindcss/vite';

// https://astro.build/config
export default defineConfig({
  site: 'https://getnetworthcalculator.com',
  trailingSlash: 'never',
  build: {
    inlineStylesheets: 'always',
    format: 'directory'
  },
  vite: {
    plugins: [tailwindcss()]
  }
});