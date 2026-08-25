// @ts-check
import { defineConfig } from 'astro/config';

import tailwindcss from '@tailwindcss/vite';

import sitemap from '@astrojs/sitemap';

// https://astro.build/config
export default defineConfig({
  site: 'https://getnetworthcalculator.com',
  trailingSlash: 'never',
  build: {
    inlineStylesheets: 'always',
    format: 'file'
  },
  vite: {
    plugins: [tailwindcss()]
  },

  integrations: [sitemap()]
});