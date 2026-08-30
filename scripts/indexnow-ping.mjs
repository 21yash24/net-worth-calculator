import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const HOST = 'getnetworthcalculator.com';
const KEY = '9f8e7d6c5b4a3f2e1d0c9b8a7f6e5d4c';
const KEY_LOCATION = `https://${HOST}/${KEY}.txt`;
const SITEMAP_PATH = path.resolve(__dirname, '../dist/sitemap.xml');

async function pingIndexNow() {
  try {
    console.log('📖 Reading dist/sitemap-0.xml to extract URLs...');
    if (!fs.existsSync(SITEMAP_PATH)) {
      console.error('❌ Sitemap not found at', SITEMAP_PATH, '. Run npm run build first.');
      return;
    }
    const sitemapContent = fs.readFileSync(SITEMAP_PATH, 'utf-8');
    const urlMatches = [...sitemapContent.matchAll(/<loc>(https:\/\/[^<]+)<\/loc>/g)];
    const urls = urlMatches.map(m => m[1]);

    console.log(`Found ${urls.length} URLs to submit to IndexNow...`);

    // IndexNow allows up to 10,000 URLs per request
    const payload = {
      host: HOST,
      key: KEY,
      keyLocation: KEY_LOCATION,
      urlList: urls
    };

    console.log('🚀 Pinging IndexNow API (Bing, Yandex, Seznam, Naver)...');
    const response = await fetch('https://api.indexnow.org/indexnow', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json; charset=utf-8'
      },
      body: JSON.stringify(payload)
    });

    if (response.ok || response.status === 200 || response.status === 202) {
      console.log(`✅ SUCCESS! IndexNow acknowledged ${urls.length} URLs (Status: ${response.status}).`);
    } else {
      const text = await response.text();
      console.log(`⚠️ IndexNow response: Status ${response.status} - ${text}`);
    }
  } catch (error) {
    console.error('❌ Error pinging IndexNow:', error.message);
  }
}

pingIndexNow();
