import fs from 'fs';
import path from 'path';

const DOMAIN = 'https://getnetworthcalculator.com';
const DIST_DIR = path.resolve('dist');

function getAllHtmlFiles(dir, fileList = []) {
  if (!fs.existsSync(dir)) return fileList;
  const files = fs.readdirSync(dir);
  for (const file of files) {
    const fullPath = path.join(dir, file);
    if (fs.statSync(fullPath).isDirectory()) {
      getAllHtmlFiles(fullPath, fileList);
    } else if (file === 'index.html') {
      const relPath = path.relative(DIST_DIR, path.dirname(fullPath));
      const route = relPath === '' ? '/' : `/${relPath.replace(/\\/g, '/')}`;
      if (!route.startsWith('/404')) {
        fileList.push(route);
      }
    }
  }
  return fileList;
}

const routes = getAllHtmlFiles(DIST_DIR);
const today = new Date().toISOString().split('T')[0];

console.log(`Generating sitemap for ${routes.length} routes...`);

let xml = `<?xml version="1.0" encoding="UTF-8"?>\n`;
xml += `<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n`;

for (const r of routes) {
  const url = r === '/' ? `${DOMAIN}/` : `${DOMAIN}${r}`;
  const priority = r === '/' ? '1.0' : (r.split('/').length <= 2 ? '0.8' : '0.6');
  const changefreq = r === '/' ? 'daily' : 'weekly';

  xml += `  <url>\n`;
  xml += `    <loc>${url}</loc>\n`;
  xml += `    <lastmod>${today}</lastmod>\n`;
  xml += `    <changefreq>${changefreq}</changefreq>\n`;
  xml += `    <priority>${priority}</priority>\n`;
  xml += `  </url>\n`;
}

xml += `</urlset>\n`;

fs.writeFileSync(path.join(DIST_DIR, 'sitemap.xml'), xml, 'utf-8');
// Also write to public so it is included in future builds
fs.writeFileSync(path.resolve('public/sitemap.xml'), xml, 'utf-8');

console.log(`✅ Generated dist/sitemap.xml and public/sitemap.xml with ${routes.length} URLs!`);
