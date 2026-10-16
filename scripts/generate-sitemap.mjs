import fs from 'fs';
import path from 'path';

const DOMAIN = 'https://getnetworthcalculator.com';
const DIST_DIR = path.resolve('dist');

function getAllHtmlRoutes(dir, fileList = []) {
  if (!fs.existsSync(dir)) return fileList;
  const files = fs.readdirSync(dir);
  for (const file of files) {
    const fullPath = path.join(dir, file);
    if (fs.statSync(fullPath).isDirectory()) {
      getAllHtmlRoutes(fullPath, fileList);
    } else if (file.endsWith('.html')) {
      if (file === '404.html') continue;
      const relPath = path.relative(DIST_DIR, fullPath).replace(/\\/g, '/');
      let route;
      if (relPath === 'index.html') {
        route = '/';
      } else if (relPath.endsWith('/index.html')) {
        route = '/' + relPath.replace(/\/index\.html$/, '');
      } else {
        route = '/' + relPath.replace(/\.html$/, '');
      }
      fileList.push(route);
    }
  }
  return fileList;
}

const rawRoutes = getAllHtmlRoutes(DIST_DIR);
// Remove duplicates and sort, ensuring root '/' is always first
const uniqueRoutes = Array.from(new Set(rawRoutes)).sort((a, b) => {
  if (a === '/') return -1;
  if (b === '/') return 1;
  return a.localeCompare(b);
});

const today = new Date().toISOString().split('T')[0];

console.log(`Generating sitemap for ${uniqueRoutes.length} clean routes...`);

// 1. Generate primary sitemap.xml with <urlset>
let sitemapXml = `<?xml version="1.0" encoding="UTF-8"?>\n`;
sitemapXml += `<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n`;

for (const r of uniqueRoutes) {
  const url = r === '/' ? `${DOMAIN}/` : `${DOMAIN}${r}`;
  const priority = r === '/' ? '1.0' : (r.split('/').length <= 2 ? '0.8' : '0.6');
  const changefreq = r === '/' ? 'daily' : 'weekly';

  sitemapXml += `  <url>\n`;
  sitemapXml += `    <loc>${url}</loc>\n`;
  sitemapXml += `    <lastmod>${today}</lastmod>\n`;
  sitemapXml += `    <changefreq>${changefreq}</changefreq>\n`;
  sitemapXml += `    <priority>${priority}</priority>\n`;
  sitemapXml += `  </url>\n`;
}

sitemapXml += `</urlset>\n`;

// 2. Generate valid sitemap-index.xml with <sitemapindex>
let sitemapIndexXml = `<?xml version="1.0" encoding="UTF-8"?>\n`;
sitemapIndexXml += `<sitemapindex xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n`;
sitemapIndexXml += `  <sitemap>\n`;
sitemapIndexXml += `    <loc>${DOMAIN}/sitemap.xml</loc>\n`;
sitemapIndexXml += `    <lastmod>${today}</lastmod>\n`;
sitemapIndexXml += `  </sitemap>\n`;
sitemapIndexXml += `</sitemapindex>\n`;

// Write to dist/ and public/
fs.writeFileSync(path.join(DIST_DIR, 'sitemap.xml'), sitemapXml, 'utf-8');
fs.writeFileSync(path.join(DIST_DIR, 'sitemap-index.xml'), sitemapIndexXml, 'utf-8');
fs.writeFileSync(path.resolve('public/sitemap.xml'), sitemapXml, 'utf-8');
fs.writeFileSync(path.resolve('public/sitemap-index.xml'), sitemapIndexXml, 'utf-8');

console.log(`✅ Successfully generated sitemap.xml (${uniqueRoutes.length} URLs) and valid sitemap-index.xml!`);
