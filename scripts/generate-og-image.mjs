import sharp from 'sharp';
import fs from 'fs';
import path from 'path';

const svg = `
<svg width="1200" height="630" viewBox="0 0 1200 630" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <linearGradient id="bg" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#090D16" />
      <stop offset="50%" stop-color="#0D1527" />
      <stop offset="100%" stop-color="#050811" />
    </linearGradient>
    <linearGradient id="emeraldGrad" x1="0%" y1="0%" x2="100%" y2="0%">
      <stop offset="0%" stop-color="#10B981" />
      <stop offset="100%" stop-color="#059669" />
    </linearGradient>
    <linearGradient id="accentGrad" x1="0%" y1="0%" x2="100%" y2="0%">
      <stop offset="0%" stop-color="#34D399" />
      <stop offset="50%" stop-color="#10B981" />
      <stop offset="100%" stop-color="#6EE7B7" />
    </linearGradient>
    <linearGradient id="curveFill" x1="0%" y1="0%" x2="0%" y2="100%">
      <stop offset="0%" stop-color="#10B981" stop-opacity="0.35" />
      <stop offset="100%" stop-color="#10B981" stop-opacity="0.0" />
    </linearGradient>
    <filter id="glow" x="-20%" y="-20%" width="140%" height="140%">
      <feGaussianBlur stdDeviation="16" result="blur" />
      <feComposite in="SourceGraphic" in2="blur" operator="over" />
    </filter>
  </defs>

  <!-- Background -->
  <rect width="1200" height="630" fill="url(#bg)" />

  <!-- Subtle ambient decorative glow circles -->
  <circle cx="1050" cy="120" r="320" fill="#10B981" opacity="0.08" filter="url(#glow)" />
  <circle cx="150" cy="520" r="280" fill="#0284C7" opacity="0.06" filter="url(#glow)" />

  <!-- Outer Border Frame -->
  <rect x="24" y="24" width="1152" height="582" rx="28" fill="none" stroke="#1E293B" stroke-width="2" />

  <!-- Top Badge -->
  <g transform="translate(80, 80)">
    <rect width="360" height="42" rx="21" fill="#064E3B" stroke="#059669" stroke-width="1.5" opacity="0.9" />
    <text x="24" y="26" font-family="-apple-system, BlinkMacSystemFont, 'Inter', 'Segoe UI', sans-serif" font-size="15" font-weight="700" fill="#A7F3D0">
      🇺🇸 Official US Federal Reserve Data (2026)
    </text>
  </g>

  <!-- Main Headline -->
  <text x="80" y="195" font-family="-apple-system, BlinkMacSystemFont, 'Inter', 'Segoe UI', sans-serif" font-size="56" font-weight="900" fill="#FFFFFF" letter-spacing="-1.5">
    Net Worth Percentile
  </text>
  <text x="80" y="260" font-family="-apple-system, BlinkMacSystemFont, 'Inter', 'Segoe UI', sans-serif" font-size="56" font-weight="900" fill="url(#accentGrad)" letter-spacing="-1.5">
    Calculator by Age
  </text>

  <!-- Subheadline -->
  <text x="80" y="315" font-family="-apple-system, BlinkMacSystemFont, 'Inter', 'Segoe UI', sans-serif" font-size="22" font-weight="500" fill="#94A3B8">
    Are you in the Top 10% for your age? Check your rank across 50 US states.
  </text>

  <!-- Bell Curve Graphic Card (Right Side) -->
  <g transform="translate(680, 80)">
    <!-- Card Base -->
    <rect width="440" height="360" rx="24" fill="#0F172A" stroke="#334155" stroke-width="1.5" />
    
    <!-- Card Title -->
    <text x="28" y="44" font-family="-apple-system, BlinkMacSystemFont, 'Inter', sans-serif" font-size="14" font-weight="700" fill="#64748B" letter-spacing="1">
      US WEALTH DISTRIBUTION CURVE
    </text>

    <!-- Curve Area & Line -->
    <g transform="translate(20, 70)">
      <!-- Filled Area under Curve -->
      <path d="M 0,180 Q 90,180 150,130 T 200,20 T 250,130 Q 310,180 400,180 L 400,185 L 0,185 Z" fill="url(#curveFill)" />
      <!-- Stroke Path -->
      <path d="M 0,180 Q 90,180 150,130 T 200,20 T 250,130 Q 310,180 400,180" fill="none" stroke="#10B981" stroke-width="4" />
      
      <!-- Top 10% Highlight Line -->
      <line x1="285" y1="20" x2="285" y2="180" stroke="#34D399" stroke-width="2" stroke-dasharray="4,4" />
      <circle cx="285" cy="88" r="8" fill="#10B981" stroke="#FFFFFF" stroke-width="3" />
      
      <!-- Top 10% Badge -->
      <rect x="250" y="25" width="110" height="32" rx="8" fill="#064E3B" stroke="#34D399" stroke-width="1" />
      <text x="264" y="46" font-family="-apple-system, BlinkMacSystemFont, 'Inter', sans-serif" font-size="13" font-weight="800" fill="#34D399">Top 10% Marker</text>
    </g>

    <!-- Stat Pillars in Card -->
    <g transform="translate(28, 280)">
      <rect width="115" height="58" rx="12" fill="#1E293B" />
      <text x="14" y="24" font-family="-apple-system, BlinkMacSystemFont, 'Inter', sans-serif" font-size="11" font-weight="600" fill="#94A3B8">Median (50th)</text>
      <text x="14" y="47" font-family="-apple-system, BlinkMacSystemFont, 'Inter', sans-serif" font-size="16" font-weight="800" fill="#F8FAFC">$135,600</text>
    </g>
    <g transform="translate(155, 280)">
      <rect width="115" height="58" rx="12" fill="#1E293B" />
      <text x="14" y="24" font-family="-apple-system, BlinkMacSystemFont, 'Inter', sans-serif" font-size="11" font-weight="600" fill="#94A3B8">Top 10% (90th)</text>
      <text x="14" y="47" font-family="-apple-system, BlinkMacSystemFont, 'Inter', sans-serif" font-size="16" font-weight="800" fill="#34D399">$980,000</text>
    </g>
    <g transform="translate(282, 280)">
      <rect width="130" height="58" rx="12" fill="#1E293B" />
      <text x="14" y="24" font-family="-apple-system, BlinkMacSystemFont, 'Inter', sans-serif" font-size="11" font-weight="600" fill="#94A3B8">Top 1% (99th)</text>
      <text x="14" y="47" font-family="-apple-system, BlinkMacSystemFont, 'Inter', sans-serif" font-size="16" font-weight="800" fill="#FBBF24">$4,500,000</text>
    </g>
  </g>

  <!-- Bottom Badges / Trust Markers -->
  <g transform="translate(80, 480)">
    <g transform="translate(0, 0)">
      <circle cx="16" cy="16" r="16" fill="#1E293B" />
      <text x="9" y="22" font-size="14">⚡</text>
      <text x="44" y="22" font-family="-apple-system, BlinkMacSystemFont, 'Inter', sans-serif" font-size="16" font-weight="700" fill="#F1F5F9">Instant &amp; 100% Client-Side Private</text>
    </g>
    <g transform="translate(360, 0)">
      <circle cx="16" cy="16" r="16" fill="#1E293B" />
      <text x="9" y="22" font-size="14">📊</text>
      <text x="44" y="22" font-family="-apple-system, BlinkMacSystemFont, 'Inter', sans-serif" font-size="16" font-weight="700" fill="#F1F5F9">Compare by Age, State &amp; Career</text>
    </g>
    <g transform="translate(720, 0)">
      <circle cx="16" cy="16" r="16" fill="#1E293B" />
      <text x="9" y="22" font-size="14">🌐</text>
      <text x="44" y="22" font-family="-apple-system, BlinkMacSystemFont, 'Inter', sans-serif" font-size="16" font-weight="700" fill="#34D399">GetNetWorthCalculator.com</text>
    </g>
  </g>
</svg>
`;

async function generate() {
  const publicPath = path.resolve('public/og-image.png');
  const distPath = path.resolve('dist/og-image.png');

  const buffer = Buffer.from(svg);
  await sharp(buffer)
    .png({ quality: 95, compressionLevel: 8 })
    .toFile(publicPath);

  if (fs.existsSync('dist')) {
    fs.copyFileSync(publicPath, distPath);
  }

  console.log(`✅ Successfully generated 1200x630 og-image.png at ${publicPath}`);
}

generate().catch(err => {
  console.error("Error generating OG image:", err);
  process.exit(1);
});
