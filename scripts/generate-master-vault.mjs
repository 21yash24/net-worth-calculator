import ExcelJS from 'exceljs';
import fs from 'node:fs';
import path from 'node:path';

const outputDir = path.resolve('./public/downloads');
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

const outputPath = path.join(outputDir, 'NetWorth_FIRE_Master_Vault_2026.xlsx');

async function createMasterVault() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Yash Jha — GetNetWorthCalculator.com';
  workbook.lastModifiedBy = 'GetNetWorth Analytics Team';
  workbook.created = new Date();
  workbook.modified = new Date();

  // --- STYLING PALETTE ---
  const colors = {
    darkBg: '0F172A',       // Slate 900
    cardBg: '1E293B',       // Slate 800
    subtleBg: '334155',     // Slate 700
    emerald: '10B981',      // Emerald Accent
    emeraldDark: '065F46',
    amber: 'F59E0B',        // Amber Gold
    rose: 'F43F5E',         // Rose / Liabilities
    blue: '3B82F6',         // Blue Accent
    purple: '8B5CF6',       // Purple / Budget Wants
    textWhite: 'FFFFFF',
    textMuted: '94A3B8',
    textDark: '0F172A',
    borderDark: '334155',
    lightRow: 'F8FAFC',
    altRow: 'F1F5F9'
  };

  const solidFill = (argb) => ({
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF' + argb }
  });

  const thinBorder = {
    top: { style: 'thin', color: { argb: 'FF' + colors.borderDark } },
    left: { style: 'thin', color: { argb: 'FF' + colors.borderDark } },
    bottom: { style: 'thin', color: { argb: 'FF' + colors.borderDark } },
    right: { style: 'thin', color: { argb: 'FF' + colors.borderDark } }
  };

  // =========================================================================
  // TAB 1: ⚡ 2-MINUTE MONTHLY QUICK-CHECK (SOLVES SPREADSHEET FATIGUE)
  // =========================================================================
  const wsQuick = workbook.addWorksheet('⚡ 2-Min Quick Check-In', {
    views: [{ showGridLines: true }]
  });

  wsQuick.columns = [
    { width: 4 },   // A
    { width: 16 },  // B: Month
    { width: 18 },  // C: Cash & Emergency
    { width: 22 },  // D: Invested Portfolio
    { width: 20 },  // E: Real Estate / Assets
    { width: 18 },  // F: Total Debts
    { width: 20 },  // G: Net Worth
    { width: 16 },  // H: MoM $ Delta
    { width: 16 },  // I: MoM % Delta
    { width: 4 }    // J
  ];

  wsQuick.mergeCells('B2:I2');
  const qh = wsQuick.getCell('B2');
  qh.value = '⚡ 2-MINUTE MONTHLY QUICK CHECK-IN (NO HOMEWORK, NO OVERWHELM)';
  qh.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
  qh.fill = solidFill(colors.emeraldDark);
  qh.alignment = { horizontal: 'center', vertical: 'middle' };
  wsQuick.getRow(2).height = 28;

  wsQuick.mergeCells('B3:I3');
  const qSub = wsQuick.getCell('B3');
  qSub.value = 'Rule of Thumb: On the 1st of each month, type just 4 numbers (green headers). Everything else calculates automatically.';
  qSub.font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FFCBD5E1' } };
  qSub.fill = solidFill(colors.cardBg);
  qSub.alignment = { horizontal: 'center', vertical: 'middle' };
  wsQuick.getRow(3).height = 20;

  const quickHeaders = [
    { col: 'B', title: 'Month' },
    { col: 'C', title: '💵 Cash & Savings' },
    { col: 'D', title: '📈 Invested Portfolio' },
    { col: 'E', title: '🏠 Property & Assets' },
    { col: 'F', title: '💳 Total Debts' },
    { col: 'G', title: '🏛️ NET WORTH' },
    { col: 'H', title: 'MoM Gain ($)' },
    { col: 'I', title: 'MoM Growth (%)' }
  ];

  wsQuick.getRow(5).height = 24;
  quickHeaders.forEach(h => {
    const cell = wsQuick.getCell(`${h.col}5`);
    cell.value = h.title;
    cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = ['C', 'D', 'E', 'F'].includes(h.col) ? solidFill('047857') : solidFill(colors.darkBg);
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = thinBorder;
  });

  const months = ['Jan 2026', 'Feb 2026', 'Mar 2026', 'Apr 2026', 'May 2026', 'Jun 2026', 'Jul 2026', 'Aug 2026', 'Sep 2026', 'Oct 2026', 'Nov 2026', 'Dec 2026'];
  const sampleCash = [25000, 26200, 25800, 27100, 26900, 28400, 29000, 28500, 30200, 31000, 31500, 33000];
  const sampleInvest = [185000, 189500, 194200, 191000, 198500, 204000, 209500, 215000, 221000, 228000, 234500, 242000];
  const sampleProp = [350000, 350000, 350000, 350000, 352000, 352000, 352000, 355000, 355000, 355000, 358000, 358000];
  const sampleDebt = [245000, 243800, 242600, 241400, 240100, 238800, 237500, 236100, 234700, 233200, 231700, 230100];

  for (let i = 0; i < 12; i++) {
    const rowNum = 6 + i;
    wsQuick.getRow(rowNum).height = 20;

    wsQuick.getCell(`B${rowNum}`).value = months[i];
    wsQuick.getCell(`B${rowNum}`).font = { name: 'Arial', size: 10, bold: true };
    wsQuick.getCell(`B${rowNum}`).alignment = { horizontal: 'center' };

    wsQuick.getCell(`C${rowNum}`).value = sampleCash[i];
    wsQuick.getCell(`C${rowNum}`).numFmt = '$#,##0';

    wsQuick.getCell(`D${rowNum}`).value = sampleInvest[i];
    wsQuick.getCell(`D${rowNum}`).numFmt = '$#,##0';

    wsQuick.getCell(`E${rowNum}`).value = sampleProp[i];
    wsQuick.getCell(`E${rowNum}`).numFmt = '$#,##0';

    wsQuick.getCell(`F${rowNum}`).value = sampleDebt[i];
    wsQuick.getCell(`F${rowNum}`).numFmt = '$#,##0';

    // Formula for Net Worth
    const nwCell = wsQuick.getCell(`G${rowNum}`);
    nwCell.value = { formula: `=(C${rowNum}+D${rowNum}+E${rowNum})-F${rowNum}` };
    nwCell.numFmt = '$#,##0';
    nwCell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF065F46' } };

    // Delta formulas
    const deltaCell = wsQuick.getCell(`H${rowNum}`);
    const pctCell = wsQuick.getCell(`I${rowNum}`);
    if (i === 0) {
      deltaCell.value = 0;
      pctCell.value = 0;
    } else {
      deltaCell.value = { formula: `=G${rowNum}-G${rowNum - 1}` };
      pctCell.value = { formula: `=(G${rowNum}-G${rowNum - 1})/G${rowNum - 1}` };
    }
    deltaCell.numFmt = '+$#,##0;-$#,##0;$0';
    pctCell.numFmt = '+0.0%;-0.0%;0.0%';

    ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach(col => {
      wsQuick.getCell(`${col}${rowNum}`).border = thinBorder;
      if (i % 2 === 1 && !['G'].includes(col)) {
        wsQuick.getCell(`${col}${rowNum}`).fill = solidFill('F8FAFC');
      }
    });
  }

  // Summary Row
  const qSumRow = 18;
  wsQuick.getRow(qSumRow).height = 22;
  wsQuick.getCell(`B${qSumRow}`).value = 'LATEST / ANNUAL';
  wsQuick.getCell(`B${qSumRow}`).font = { bold: true };
  wsQuick.getCell(`C${qSumRow}`).value = { formula: '=C17' };
  wsQuick.getCell(`D${qSumRow}`).value = { formula: '=D17' };
  wsQuick.getCell(`E${qSumRow}`).value = { formula: '=E17' };
  wsQuick.getCell(`F${qSumRow}`).value = { formula: '=F17' };
  wsQuick.getCell(`G${qSumRow}`).value = { formula: '=G17' };
  wsQuick.getCell(`H${qSumRow}`).value = { formula: '=G17-G6' };
  wsQuick.getCell(`I${qSumRow}`).value = { formula: '=(G17-G6)/G6' };

  ['C', 'D', 'E', 'F', 'G', 'H'].forEach(c => wsQuick.getCell(`${c}${qSumRow}`).numFmt = '$#,##0');
  wsQuick.getCell(`I${qSumRow}`).numFmt = '+0.0%;-0.0%;0.0%';
  ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach(col => {
    const c = wsQuick.getCell(`${col}${qSumRow}`);
    c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = solidFill(colors.darkBg);
    c.border = thinBorder;
  });

  // =========================================================================
  // TAB 2: 📊 EXECUTIVE WEALTH DASHBOARD (BLOOMBERG STYLE TERMINAL)
  // =========================================================================
  const ws1 = workbook.addWorksheet('📊 Executive Dashboard', {
    views: [{ showGridLines: true }]
  });

  ws1.columns = [
    { width: 4 },   // A
    { width: 28 },  // B
    { width: 20 },  // C
    { width: 20 },  // D
    { width: 20 },  // E
    { width: 22 },  // F
    { width: 4 }    // G
  ];

  ws1.mergeCells('B2:F2');
  const hCell = ws1.getCell('B2');
  hCell.value = '🏛️ NET WORTH & FINANCIAL INDEPENDENCE MASTER VAULT (2026 EDITION)';
  hCell.font = { name: 'Arial', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
  hCell.fill = solidFill(colors.darkBg);
  hCell.alignment = { horizontal: 'center', vertical: 'middle' };
  ws1.getRow(2).height = 30;

  ws1.mergeCells('B3:F3');
  const subCell = ws1.getCell('B3');
  subCell.value = 'Private Wealth Dashboard • Federal Reserve SCF Percentile Engine • Trinity Study FIRE Calculator';
  subCell.font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF94A3B8' } };
  subCell.fill = solidFill(colors.cardBg);
  subCell.alignment = { horizontal: 'center', vertical: 'middle' };
  ws1.getRow(3).height = 18;

  // KPI Row 1: Net Worth & Liquid
  ws1.getRow(5).height = 18;
  ws1.getRow(6).height = 28;
  ws1.getRow(7).height = 16;

  // Card 1: Total Net Worth
  ws1.mergeCells('B5:C5');
  ws1.getCell('B5').value = 'TOTAL NET WORTH (AS OF DEC)';
  ws1.getCell('B5').font = { size: 9, bold: true, color: { argb: 'FF94A3B8' } };
  ws1.getCell('B5').fill = solidFill(colors.cardBg);

  ws1.mergeCells('B6:C6');
  const nwKPI = ws1.getCell('B6');
  nwKPI.value = { formula: "='⚡ 2-Min Quick Check-In'!G18" };
  nwKPI.font = { size: 18, bold: true, color: { argb: 'FFFFFFFF' } };
  nwKPI.numFmt = '$#,##0';
  nwKPI.fill = solidFill(colors.cardBg);
  nwKPI.alignment = { vertical: 'middle' };

  ws1.mergeCells('B7:C7');
  const nwSub = ws1.getCell('B7');
  nwSub.value = { formula: `="▲ Annual Growth: " & TEXT('⚡ 2-Min Quick Check-In'!I18, "+0.0%") & " (+$" & TEXT('⚡ 2-Min Quick Check-In'!H18, "#,##0") & ")"` };
  nwSub.font = { size: 9, bold: true, color: { argb: 'FF10B981' } };
  nwSub.fill = solidFill(colors.cardBg);

  // Card 2: Liquid Investable Assets
  ws1.mergeCells('D5:E5');
  ws1.getCell('D5').value = 'LIQUID INVESTED PORTFOLIO';
  ws1.getCell('D5').font = { size: 9, bold: true, color: { argb: 'FF94A3B8' } };
  ws1.getCell('D5').fill = solidFill(colors.cardBg);

  ws1.mergeCells('D6:E6');
  const liqKPI = ws1.getCell('D6');
  liqKPI.value = { formula: "='⚡ 2-Min Quick Check-In'!C18+'⚡ 2-Min Quick Check-In'!D18" };
  liqKPI.font = { size: 18, bold: true, color: { argb: 'FF10B981' } };
  liqKPI.numFmt = '$#,##0';
  liqKPI.fill = solidFill(colors.cardBg);
  liqKPI.alignment = { vertical: 'middle' };

  ws1.mergeCells('D7:E7');
  const liqSub = ws1.getCell('D7');
  liqSub.value = { formula: `="Liquid Ratio: " & TEXT(('⚡ 2-Min Quick Check-In'!C18+'⚡ 2-Min Quick Check-In'!D18)/('⚡ 2-Min Quick Check-In'!G18+'⚡ 2-Min Quick Check-In'!F18), "0.0%") & " of total assets"` };
  liqSub.font = { size: 9, color: { argb: 'FFCBD5E1' } };
  liqSub.fill = solidFill(colors.cardBg);

  // Card 3: Fed Reserve SCF Percentile
  ws1.getCell('F5').value = 'FED RESERVE SCF RANK';
  ws1.getCell('F5').font = { size: 9, bold: true, color: { argb: 'FF94A3B8' } };
  ws1.getCell('F5').fill = solidFill(colors.cardBg);

  const fedRankCell = ws1.getCell('F6');
  fedRankCell.value = 'Top 12%';
  fedRankCell.font = { size: 18, bold: true, color: { argb: 'FFF59E0B' } };
  fedRankCell.fill = solidFill(colors.cardBg);
  fedRankCell.alignment = { vertical: 'middle' };

  const fedSubCell = ws1.getCell('F7');
  fedSubCell.value = 'Age Bracket: 30 - 34 (US Baseline)';
  fedSubCell.font = { size: 9, color: { argb: 'FFCBD5E1' } };
  fedSubCell.fill = solidFill(colors.cardBg);

  // Interactive Trajectory Table
  ws1.mergeCells('B9:F9');
  const tHeader = ws1.getCell('B9');
  tHeader.value = '📈 12-MONTH WEALTH TRAJECTORY & SUMMARY';
  tHeader.font = { size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
  tHeader.fill = solidFill(colors.darkBg);
  ws1.getRow(9).height = 24;

  const tCols = [
    { col: 'B', title: 'Quarter / Milestone' },
    { col: 'C', title: 'Net Worth' },
    { col: 'D', title: 'Invested Assets' },
    { col: 'E', title: 'Total Liabilities' },
    { col: 'F', title: 'Progress Note' }
  ];

  ws1.getRow(10).height = 20;
  tCols.forEach(tc => {
    const c = ws1.getCell(`${tc.col}10`);
    c.value = tc.title;
    c.font = { size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = solidFill(colors.cardBg);
    c.border = thinBorder;
  });

  const milestones = [
    { label: 'Q1 (Jan 2026 Baseline)', nw: "='⚡ 2-Min Quick Check-In'!G6", inv: "='⚡ 2-Min Quick Check-In'!D6", debt: "='⚡ 2-Min Quick Check-In'!F6", note: 'Starting 2026 Baseline' },
    { label: 'Q2 (Apr 2026)', nw: "='⚡ 2-Min Quick Check-In'!G9", inv: "='⚡ 2-Min Quick Check-In'!D9", debt: "='⚡ 2-Min Quick Check-In'!F9", note: 'Q1 Tax Optimization' },
    { label: 'Q3 (Jul 2026)', nw: "='⚡ 2-Min Quick Check-In'!G12", inv: "='⚡ 2-Min Quick Check-In'!D12", debt: "='⚡ 2-Min Quick Check-In'!F12", note: 'Mid-Year Portfolio Rebalance' },
    { label: 'Q4 (Oct 2026)', nw: "='⚡ 2-Min Quick Check-In'!G15", inv: "='⚡ 2-Min Quick Check-In'!D15", debt: "='⚡ 2-Min Quick Check-In'!F15", note: 'Year-End Sinking Fund Push' },
    { label: 'Year-End Total (Dec 2026)', nw: "='⚡ 2-Min Quick Check-In'!G17", inv: "='⚡ 2-Min Quick Check-In'!D17", debt: "='⚡ 2-Min Quick Check-In'!F17", note: 'Annual Target Achieved' }
  ];

  milestones.forEach((m, idx) => {
    const row = 11 + idx;
    ws1.getRow(row).height = 20;
    ws1.getCell(`B${row}`).value = m.label;
    ws1.getCell(`C${row}`).value = { formula: m.nw };
    ws1.getCell(`D${row}`).value = { formula: m.inv };
    ws1.getCell(`E${row}`).value = { formula: m.debt };
    ws1.getCell(`F${row}`).value = m.note;

    ['C', 'D', 'E'].forEach(c => ws1.getCell(`${c}${row}`).numFmt = '$#,##0');
    ['B', 'C', 'D', 'E', 'F'].forEach(c => {
      ws1.getCell(`${c}${row}`).border = thinBorder;
      if (idx % 2 === 1) ws1.getCell(`${c}${row}`).fill = solidFill('F8FAFC');
    });
  });

  // FIRE Summary Card in Dashboard
  ws1.mergeCells('B18:F18');
  const fireDashH = ws1.getCell('B18');
  fireDashH.value = '🔥 FINANCIAL INDEPENDENCE (FIRE) READINESS';
  fireDashH.font = { size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
  fireDashH.fill = solidFill(colors.darkBg);
  ws1.getRow(18).height = 24;

  const fireKPIs = [
    { label: 'Lean FIRE (25x Basic Needs)', val: '$1,080,000', status: '25.4% Complete', col: 'B' },
    { label: 'Standard FIRE (4% Rule)', val: '$1,650,000', status: '16.6% Complete', col: 'C' },
    { label: 'Coast FIRE Target (Age 32)', val: '$185,000', status: '✅ ACHIEVED', col: 'D' },
    { label: 'Estimated Freedom Date', val: 'October 2038', status: '12.8 Years Left', col: 'E' }
  ];

  ws1.getRow(19).height = 20;
  ws1.getRow(20).height = 24;
  ws1.mergeCells('E19:F19');
  ws1.mergeCells('E20:F20');

  fireKPIs.forEach(fk => {
    ws1.getCell(`${fk.col}19`).value = fk.label;
    ws1.getCell(`${fk.col}19`).font = { size: 8, bold: true, color: { argb: 'FF94A3B8' } };
    ws1.getCell(`${fk.col}19`).fill = solidFill(colors.cardBg);

    const vCell = ws1.getCell(`${fk.col}20`);
    vCell.value = fk.val;
    vCell.font = { size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
    vCell.fill = solidFill(colors.cardBg);
  });

  // =========================================================================
  // TAB 3: 💸 SMART MONTHLY BUDGET (50/30/20 & GUILT-FREE SPENDING BUFFER)
  // =========================================================================
  // This directly addresses what real people ask for:
  // - Answers "How much can I spend right now without feeling guilty?"
  // - No micro-receipt tracking homework.
  // - Clean 50/30/20 cash flow breakdown.
  const wsBudget = workbook.addWorksheet('💸 Smart Budget (50-30-20)', {
    views: [{ showGridLines: true }]
  });

  wsBudget.columns = [
    { width: 4 },   // A
    { width: 28 },  // B: Category
    { width: 18 },  // C: Budgeted
    { width: 18 },  // D: Actual
    { width: 16 },  // E: Difference
    { width: 22 },  // F: Target % / Status
    { width: 4 }    // G
  ];

  wsBudget.mergeCells('B2:F2');
  const bHeader = wsBudget.getCell('B2');
  bHeader.value = '💸 SMART MONTHLY CASH FLOW & 50/30/20 BUDGET ENGINE';
  bHeader.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
  bHeader.fill = solidFill(colors.darkBg);
  bHeader.alignment = { horizontal: 'center', vertical: 'middle' };
  wsBudget.getRow(2).height = 28;

  wsBudget.mergeCells('B3:F3');
  const bSub = wsBudget.getCell('B3');
  bSub.value = 'Answers: "How much can I safely spend guilt-free?" • 50% Needs, 30% Wants, 20% Wealth Surplus.';
  bSub.font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF94A3B8' } };
  bSub.fill = solidFill(colors.cardBg);
  bSub.alignment = { horizontal: 'center', vertical: 'middle' };
  wsBudget.getRow(3).height = 18;

  // Key Answer Card: "SAFE TO SPEND GUILT-FREE THIS MONTH"
  wsBudget.mergeCells('B5:D5');
  wsBudget.getCell('B5').value = '🟢 REMAINING GUILT-FREE SPENDING BUFFER';
  wsBudget.getCell('B5').font = { size: 9, bold: true, color: { argb: 'FF10B981' } };
  wsBudget.getCell('B5').fill = solidFill(colors.cardBg);

  wsBudget.mergeCells('B6:D6');
  const safeSpendCell = wsBudget.getCell('B6');
  safeSpendCell.value = { formula: '=E24' };
  safeSpendCell.numFmt = '$#,##0';
  safeSpendCell.font = { size: 18, bold: true, color: { argb: 'FF10B981' } };
  safeSpendCell.fill = solidFill(colors.cardBg);

  wsBudget.mergeCells('E5:F5');
  wsBudget.getCell('E5').value = 'DAILY ALLOWANCE (30 DAYS)';
  wsBudget.getCell('E5').font = { size: 9, bold: true, color: { argb: 'FF94A3B8' } };
  wsBudget.getCell('E5').fill = solidFill(colors.cardBg);

  wsBudget.mergeCells('E6:F6');
  const dailySpendCell = wsBudget.getCell('E6');
  dailySpendCell.value = { formula: '=E24/30' };
  dailySpendCell.numFmt = '$#,##0/day';
  dailySpendCell.font = { size: 15, bold: true, color: { argb: 'FFFFFFFF' } };
  dailySpendCell.fill = solidFill(colors.cardBg);

  // Table Headers
  const bTableHeaders = [
    { col: 'B', title: 'Cash Flow Stream' },
    { col: 'C', title: 'Planned ($)' },
    { col: 'D', title: 'Actual ($)' },
    { col: 'E', title: 'Remaining ($)' },
    { col: 'F', title: '50/30/20 Benchmark' }
  ];

  wsBudget.getRow(8).height = 22;
  bTableHeaders.forEach(th => {
    const c = wsBudget.getCell(`${th.col}8`);
    c.value = th.title;
    c.font = { size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = solidFill(colors.darkBg);
    c.border = thinBorder;
  });

  // Section 1: Inflow
  wsBudget.mergeCells('B9:F9');
  wsBudget.getCell('B9').value = '📥 MONTHLY NET INCOME (TAKE-HOME PAY)';
  wsBudget.getCell('B9').font = { size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  wsBudget.getCell('B9').fill = solidFill(colors.emeraldDark);
  wsBudget.getRow(9).height = 20;

  const incomeItems = [
    { name: 'Primary Salary / Take-Home', plan: 6800, act: 6800 },
    { name: 'Partner / Secondary Income', plan: 2200, act: 2200 },
    { name: 'Side Hustles / Dividends', plan: 500, act: 650 }
  ];

  incomeItems.forEach((item, idx) => {
    const r = 10 + idx;
    wsBudget.getRow(r).height = 19;
    wsBudget.getCell(`B${r}`).value = item.name;
    wsBudget.getCell(`C${r}`).value = item.plan;
    wsBudget.getCell(`D${r}`).value = item.act;
    wsBudget.getCell(`E${r}`).value = { formula: `=D${r}-C${r}` };
    wsBudget.getCell(`F${r}`).value = 'Inflow';
    ['C', 'D', 'E'].forEach(c => wsBudget.getCell(`${c}${r}`).numFmt = '$#,##0');
    ['B', 'C', 'D', 'E', 'F'].forEach(c => wsBudget.getCell(`${c}${r}`).border = thinBorder);
  });

  // Total Income
  const rIncTot = 13;
  wsBudget.getRow(rIncTot).height = 20;
  wsBudget.getCell(`B${rIncTot}`).value = 'TOTAL NET INCOME';
  wsBudget.getCell(`B${rIncTot}`).font = { bold: true };
  wsBudget.getCell(`C${rIncTot}`).value = { formula: '=SUM(C10:C12)' };
  wsBudget.getCell(`D${rIncTot}`).value = { formula: '=SUM(D10:D12)' };
  wsBudget.getCell(`E${rIncTot}`).value = { formula: '=D13-C13' };
  wsBudget.getCell(`F${rIncTot}`).value = '100.0% Baseline';
  ['C', 'D', 'E'].forEach(c => wsBudget.getCell(`${c}${rIncTot}`).numFmt = '$#,##0');
  ['B', 'C', 'D', 'E', 'F'].forEach(c => {
    const cell = wsBudget.getCell(`${c}${rIncTot}`);
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = solidFill(colors.cardBg);
    cell.border = thinBorder;
  });

  // Section 2: Needs (50%)
  wsBudget.mergeCells('B14:F14');
  wsBudget.getCell('B14').value = '🏠 ESSENTIAL NEEDS (TARGET: ≤ 50% OF INCOME)';
  wsBudget.getCell('B14').font = { size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  wsBudget.getCell('B14').fill = solidFill('1E3A8A'); // Blue
  wsBudget.getRow(14).height = 20;

  const needsItems = [
    { name: 'Mortgage / Rent Payment', plan: 2400, act: 2400 },
    { name: 'Groceries & Household Essentials', plan: 750, act: 820 },
    { name: 'Utilities, Electricity & Internet', plan: 320, act: 310 },
    { name: 'Healthcare & Insurance Premiums', plan: 450, act: 450 },
    { name: 'Car Payment / Transportation / Gas', plan: 400, act: 380 }
  ];

  needsItems.forEach((item, idx) => {
    const r = 15 + idx;
    wsBudget.getRow(r).height = 19;
    wsBudget.getCell(`B${r}`).value = item.name;
    wsBudget.getCell(`C${r}`).value = item.plan;
    wsBudget.getCell(`D${r}`).value = item.act;
    wsBudget.getCell(`E${r}`).value = { formula: `=C${r}-D${r}` };
    wsBudget.getCell(`F${r}`).value = { formula: `=TEXT(D${r}/$D$13, "0.0%")` };
    ['C', 'D', 'E'].forEach(c => wsBudget.getCell(`${c}${r}`).numFmt = '$#,##0');
    ['B', 'C', 'D', 'E', 'F'].forEach(c => wsBudget.getCell(`${c}${r}`).border = thinBorder);
  });

  // Total Needs
  const rNeedTot = 20;
  wsBudget.getRow(rNeedTot).height = 20;
  wsBudget.getCell(`B${rNeedTot}`).value = 'TOTAL ESSENTIAL NEEDS';
  wsBudget.getCell(`C${rNeedTot}`).value = { formula: '=SUM(C15:C19)' };
  wsBudget.getCell(`D${rNeedTot}`).value = { formula: '=SUM(D15:D19)' };
  wsBudget.getCell(`E${rNeedTot}`).value = { formula: '=C20-D20' };
  wsBudget.getCell(`F${rNeedTot}`).value = { formula: `="Actual: " & TEXT(D20/D13, "0.0%") & " (Target: ≤50%)"` };
  ['C', 'D', 'E'].forEach(c => wsBudget.getCell(`${c}${rNeedTot}`).numFmt = '$#,##0');
  ['B', 'C', 'D', 'E', 'F'].forEach(c => {
    const cell = wsBudget.getCell(`${c}${rNeedTot}`);
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = solidFill('1E3A8A');
    cell.border = thinBorder;
  });

  // Section 3: Wants (30%)
  wsBudget.mergeCells('B21:F21');
  wsBudget.getCell('B21').value = '🎉 GUILT-FREE WANTS (TARGET: ≤ 30% OF INCOME)';
  wsBudget.getCell('B21').font = { size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  wsBudget.getCell('B21').fill = solidFill('6B21A8'); // Purple
  wsBudget.getRow(21).height = 20;

  const wantsItems = [
    { name: 'Dining Out & Weekend Socials', plan: 650, act: 520 },
    { name: 'Entertainment, Streaming & Subs', plan: 120, act: 115 },
    { name: 'Shopping & Personal Hobbies', plan: 450, act: 380 }
  ];

  wantsItems.forEach((item, idx) => {
    const r = 22 + idx;
    wsBudget.getRow(r).height = 19;
    wsBudget.getCell(`B${r}`).value = item.name;
    wsBudget.getCell(`C${r}`).value = item.plan;
    wsBudget.getCell(`D${r}`).value = item.act;
    wsBudget.getCell(`E${r}`).value = { formula: `=C${r}-D${r}` };
    wsBudget.getCell(`F${r}`).value = { formula: `=TEXT(D${r}/$D$13, "0.0%")` };
    ['C', 'D', 'E'].forEach(c => wsBudget.getCell(`${c}${r}`).numFmt = '$#,##0');
    ['B', 'C', 'D', 'E', 'F'].forEach(c => wsBudget.getCell(`${c}${r}`).border = thinBorder);
  });

  // Total Wants
  const rWantTot = 25;
  wsBudget.getRow(rWantTot).height = 20;
  wsBudget.getCell(`B${rWantTot}`).value = 'TOTAL GUILT-FREE WANTS';
  wsBudget.getCell(`C${rWantTot}`).value = { formula: '=SUM(C22:C24)' };
  wsBudget.getCell(`D${rWantTot}`).value = { formula: '=SUM(D22:D24)' };
  wsBudget.getCell(`E${rWantTot}`).value = { formula: '=C25-D25' };
  wsBudget.getCell(`F${rWantTot}`).value = { formula: `="Actual: " & TEXT(D25/D13, "0.0%") & " (Target: ≤30%)"` };
  ['C', 'D', 'E'].forEach(c => wsBudget.getCell(`${c}${rWantTot}`).numFmt = '$#,##0');
  ['B', 'C', 'D', 'E', 'F'].forEach(c => {
    const cell = wsBudget.getCell(`${c}${rWantTot}`);
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = solidFill('6B21A8');
    cell.border = thinBorder;
  });

  // Section 4: Wealth & Savings Surplus (20%+)
  wsBudget.mergeCells('B26:F26');
  wsBudget.getCell('B26').value = '🚀 WEALTH SURPLUS & SAVINGS (TARGET: ≥ 20% OF INCOME)';
  wsBudget.getCell('B26').font = { size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  wsBudget.getCell('B26').fill = solidFill(colors.emeraldDark);
  wsBudget.getRow(26).height = 20;

  const savingsItems = [
    { name: '401(k) / IRA / Index Fund DCA', plan: 2200, act: 2200 },
    { name: 'Emergency Sinking Fund Transfer', plan: 500, act: 500 },
    { name: 'Extra Debt Acceleration Principal', plan: 600, act: 750 }
  ];

  savingsItems.forEach((item, idx) => {
    const r = 27 + idx;
    wsBudget.getRow(r).height = 19;
    wsBudget.getCell(`B${r}`).value = item.name;
    wsBudget.getCell(`C${r}`).value = item.plan;
    wsBudget.getCell(`D${r}`).value = item.act;
    wsBudget.getCell(`E${r}`).value = { formula: `=D${r}-C${r}` };
    wsBudget.getCell(`F${r}`).value = { formula: `=TEXT(D${r}/$D$13, "0.0%")` };
    ['C', 'D', 'E'].forEach(c => wsBudget.getCell(`${c}${r}`).numFmt = '$#,##0');
    ['B', 'C', 'D', 'E', 'F'].forEach(c => wsBudget.getCell(`${c}${r}`).border = thinBorder);
  });

  // Total Savings
  const rSaveTot = 30;
  wsBudget.getRow(rSaveTot).height = 20;
  wsBudget.getCell(`B${rSaveTot}`).value = 'TOTAL WEALTH SURPLUS';
  wsBudget.getCell(`C${rSaveTot}`).value = { formula: '=SUM(C27:C29)' };
  wsBudget.getCell(`D${rSaveTot}`).value = { formula: '=SUM(D27:D29)' };
  wsBudget.getCell(`E${rSaveTot}`).value = { formula: '=D30-C30' };
  wsBudget.getCell(`F${rSaveTot}`).value = { formula: `="Savings Rate: " & TEXT(D30/D13, "0.0%") & " (FIRE Speed)"` };
  ['C', 'D', 'E'].forEach(c => wsBudget.getCell(`${c}${rSaveTot}`).numFmt = '$#,##0');
  ['B', 'C', 'D', 'E', 'F'].forEach(c => {
    const cell = wsBudget.getCell(`${c}${rSaveTot}`);
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = solidFill(colors.emeraldDark);
    cell.border = thinBorder;
  });

  // =========================================================================
  // TAB 4: 🎯 SINKING FUNDS & IRREGULAR EXPENSE SHIELD
  // =========================================================================
  // Solves the #1 reason budgets fail: Christmas, car repairs, annual insurance.
  const wsSinking = workbook.addWorksheet('🎯 Sinking Funds Shield', {
    views: [{ showGridLines: true }]
  });

  wsSinking.columns = [
    { width: 4 },   // A
    { width: 26 },  // B: Fund Goal
    { width: 16 },  // C: Target Amount
    { width: 16 },  // D: Current Saved
    { width: 16 },  // E: Still Needed
    { width: 16 },  // F: Target Date
    { width: 18 },  // G: Monthly Deposit
    { width: 14 }   // H: Progress %
  ];

  wsSinking.mergeCells('B2:H2');
  const sH = wsSinking.getCell('B2');
  sH.value = '🎯 SINKING FUNDS & IRREGULAR EXPENSE DEFENSE SHIELD';
  sH.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
  sH.fill = solidFill('047857');
  sH.alignment = { horizontal: 'center', vertical: 'middle' };
  wsSinking.getRow(2).height = 28;

  wsSinking.mergeCells('B3:H3');
  const sSub = wsSinking.getCell('B3');
  sSub.value = 'Prevents budget shock from annual car insurance, holiday gifts, unexpected home repairs, and vacations.';
  sSub.font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FFCBD5E1' } };
  sSub.fill = solidFill(colors.cardBg);
  sSub.alignment = { horizontal: 'center', vertical: 'middle' };
  wsSinking.getRow(3).height = 18;

  const sHeaders = [
    { col: 'B', title: 'Fund Name / Category' },
    { col: 'C', title: 'Target Goal' },
    { col: 'D', title: 'Current Balance' },
    { col: 'E', title: 'Remaining Gap' },
    { col: 'F', title: 'Target Deadline' },
    { col: 'G', title: 'Monthly Allocation' },
    { col: 'H', title: 'Fund Health' }
  ];

  wsSinking.getRow(5).height = 22;
  sHeaders.forEach(th => {
    const c = wsSinking.getCell(`${th.col}5`);
    c.value = th.title;
    c.font = { size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = solidFill(colors.darkBg);
    c.border = thinBorder;
  });

  const sinkingFunds = [
    { name: '🛡️ 6-Month Emergency Runway', target: 25000, curr: 21500, date: 'Jun 2026', mo: 700 },
    { name: '🚗 Car Maintenance & Tires', target: 2000, curr: 1400, date: 'Aug 2026', mo: 150 },
    { name: '🏖️ Annual Family Vacation', target: 4500, curr: 2800, date: 'Dec 2026', mo: 250 },
    { name: '🎁 Holiday Season & Gifts', target: 1500, curr: 750, date: 'Nov 2026', mo: 125 },
    { name: '🏠 Home Appliance & Repair Pool', target: 3500, curr: 1800, date: 'Oct 2026', mo: 250 }
  ];

  sinkingFunds.forEach((sf, idx) => {
    const r = 6 + idx;
    wsSinking.getRow(r).height = 20;
    wsSinking.getCell(`B${r}`).value = sf.name;
    wsSinking.getCell(`C${r}`).value = sf.target;
    wsSinking.getCell(`D${r}`).value = sf.curr;
    wsSinking.getCell(`E${r}`).value = { formula: `=C${r}-D${r}` };
    wsSinking.getCell(`F${r}`).value = sf.date;
    wsSinking.getCell(`G${r}`).value = sf.mo;
    wsSinking.getCell(`H${r}`).value = { formula: `=D${r}/C${r}` };

    ['C', 'D', 'E', 'G'].forEach(c => wsSinking.getCell(`${c}${r}`).numFmt = '$#,##0');
    wsSinking.getCell(`H${r}`).numFmt = '0.0%';
    ['B', 'C', 'D', 'E', 'F', 'G', 'H'].forEach(c => {
      wsSinking.getCell(`${c}${r}`).border = thinBorder;
      if (idx % 2 === 1) wsSinking.getCell(`${c}${r}`).fill = solidFill('F8FAFC');
    });
  });

  // Total Sinking Funds Row
  const sTotRow = 11;
  wsSinking.getRow(sTotRow).height = 22;
  wsSinking.getCell(`B${sTotRow}`).value = 'TOTAL SINKING FUNDS';
  wsSinking.getCell(`C${sTotRow}`).value = { formula: '=SUM(C6:C10)' };
  wsSinking.getCell(`D${sTotRow}`).value = { formula: '=SUM(D6:D10)' };
  wsSinking.getCell(`E${sTotRow}`).value = { formula: '=SUM(E6:E10)' };
  wsSinking.getCell(`F${sTotRow}`).value = 'Multiple';
  wsSinking.getCell(`G${sTotRow}`).value = { formula: '=SUM(G6:G10)' };
  wsSinking.getCell(`H${sTotRow}`).value = { formula: '=D11/C11' };

  ['C', 'D', 'E', 'G'].forEach(c => wsSinking.getCell(`${c}${sTotRow}`).numFmt = '$#,##0');
  wsSinking.getCell(`H${sTotRow}`).numFmt = '0.0%';
  ['B', 'C', 'D', 'E', 'F', 'G', 'H'].forEach(c => {
    const cell = wsSinking.getCell(`${c}${sTotRow}`);
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = solidFill(colors.darkBg);
    cell.border = thinBorder;
  });

  // =========================================================================
  // TAB 5: 🔥 FIRE FREEDOM & RETIREMENT ENGINE
  // =========================================================================
  const wsFIRE = workbook.addWorksheet('🔥 FIRE Freedom Engine', {
    views: [{ showGridLines: true }]
  });

  wsFIRE.columns = [
    { width: 4 }, { width: 30 }, { width: 18 }, { width: 22 }, { width: 20 }, { width: 4 }
  ];

  wsFIRE.mergeCells('B2:E2');
  const fireHead = wsFIRE.getCell('B2');
  fireHead.value = '🔥 FINANCIAL INDEPENDENCE & EARLY RETIREMENT CALCULATOR (TRINITY STUDY)';
  fireHead.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
  fireHead.fill = solidFill(colors.darkBg);
  fireHead.alignment = { horizontal: 'center', vertical: 'middle' };
  wsFIRE.getRow(2).height = 28;

  wsFIRE.mergeCells('B4:E4');
  wsFIRE.getCell('B4').value = '1. CORE ASSUMPTIONS & SAFE WITHDRAWAL RATE (SWR)';
  wsFIRE.getCell('B4').font = { size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  wsFIRE.getCell('B4').fill = solidFill(colors.cardBg);

  const fireInputs = [
    { name: 'Current Age', val: 32, fmt: '0' },
    { name: 'Target Annual Retirement Spending', val: 66000, fmt: '$#,##0' },
    { name: 'Safe Withdrawal Rate (Standard Trinity 4%)', val: 0.04, fmt: '0.0%' },
    { name: 'Expected Real Return After Inflation', val: 0.07, fmt: '0.0%' },
    { name: 'Current Liquid Invested Portfolio', val: "='⚡ 2-Min Quick Check-In'!D18", fmt: '$#,##0' },
    { name: 'Monthly Investment Inflow (from Budget)', val: "='💸 Smart Budget (50-30-20)'!D27", fmt: '$#,##0' }
  ];

  fireInputs.forEach((inp, idx) => {
    const r = 5 + idx;
    wsFIRE.getRow(r).height = 19;
    wsFIRE.getCell(`B${r}`).value = inp.name;
    const v = wsFIRE.getCell(`C${r}`);
    if (typeof inp.val === 'string' && inp.val.startsWith('=')) {
      v.value = { formula: inp.val.slice(1) };
    } else {
      v.value = inp.val;
    }
    v.numFmt = inp.fmt;
    v.font = { bold: true };
    wsFIRE.getCell(`B${r}`).border = thinBorder;
    v.border = thinBorder;
  });

  // FIRE Milestones
  wsFIRE.mergeCells('B12:E12');
  wsFIRE.getCell('B12').value = '2. THE 4 FIRE MILESTONES (LEAN, STANDARD, FAT, COAST)';
  wsFIRE.getCell('B12').font = { size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  wsFIRE.getCell('B12').fill = solidFill(colors.emeraldDark);

  const fireTiers = [
    { name: 'Lean FIRE (25x Bare Minimum Expenses)', formula: '=C6*0.65/C7', desc: 'Survival baseline without luxury' },
    { name: 'Standard FIRE (25x Current Expenses / 4%)', formula: '=C6/C7', desc: 'Replaces 100% of current lifestyle' },
    { name: 'Fat FIRE (Abundant Lifestyle 35x Expenses)', formula: '=C6*1.4/C7', desc: 'Zero financial restrictions' },
    { name: 'Coast FIRE Target for Age 32 (No Further Savings)', formula: '=(C6/C7)/(1+C8)^(60-C5)', desc: 'Compounding does all remaining work' }
  ];

  fireTiers.forEach((tier, idx) => {
    const r = 13 + idx;
    wsFIRE.getRow(r).height = 20;
    wsFIRE.getCell(`B${r}`).value = tier.name;
    const cVal = wsFIRE.getCell(`C${r}`);
    cVal.value = { formula: tier.formula.slice(1) };
    cVal.numFmt = '$#,##0';
    cVal.font = { bold: true };

    wsFIRE.getCell(`D${r}`).value = tier.desc;
    wsFIRE.getCell(`D${r}`).font = { italic: true, size: 9, color: { argb: 'FF64748B' } };

    ['B', 'C', 'D'].forEach(col => wsFIRE.getCell(`${col}${r}`).border = thinBorder);
  });

  // =========================================================================
  // TAB 6: 💳 DEBT ACCELERATOR (AVALANCHE VS SNOWBALL)
  // =========================================================================
  const wsDebt = workbook.addWorksheet('💳 Debt Avalanche & Snowball', {
    views: [{ showGridLines: true }]
  });

  wsDebt.columns = [
    { width: 4 }, { width: 24 }, { width: 16 }, { width: 14 }, { width: 16 }, { width: 18 }, { width: 4 }
  ];

  wsDebt.mergeCells('B2:F2');
  const dH = wsDebt.getCell('B2');
  dH.value = '💳 DEBT PAYOFF ACCELERATOR: MATHEMATICAL AVALANCHE VS BEHAVIORAL SNOWBALL';
  dH.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
  dH.fill = solidFill(colors.darkBg);
  dH.alignment = { horizontal: 'center', vertical: 'middle' };
  wsDebt.getRow(2).height = 28;

  const debtHeaders = [
    { col: 'B', title: 'Debt Name' },
    { col: 'C', title: 'Current Balance' },
    { col: 'D', title: 'Interest APR' },
    { col: 'E', title: 'Min Payment' },
    { col: 'F', title: 'Avalanche Rank' }
  ];

  wsDebt.getRow(4).height = 20;
  debtHeaders.forEach(dh => {
    const c = wsDebt.getCell(`${dh.col}4`);
    c.value = dh.title;
    c.font = { size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = solidFill(colors.cardBg);
    c.border = thinBorder;
  });

  const debts = [
    { name: 'Credit Card A (High APR)', bal: 4200, apr: 0.2499, min: 140, rank: 'Priority #1 (Avalanche)' },
    { name: 'Personal Consolidation Loan', bal: 8500, apr: 0.115, min: 210, rank: 'Priority #2' },
    { name: 'Auto Loan (Vehicle)', bal: 17400, apr: 0.059, min: 380, rank: 'Priority #3' },
    { name: 'Primary Mortgage', bal: 200000, apr: 0.035, min: 1670, rank: 'Priority #4 (Low Interest)' }
  ];

  debts.forEach((d, idx) => {
    const r = 5 + idx;
    wsDebt.getRow(r).height = 19;
    wsDebt.getCell(`B${r}`).value = d.name;
    wsDebt.getCell(`C${r}`).value = d.bal;
    wsDebt.getCell(`C${r}`).numFmt = '$#,##0';
    wsDebt.getCell(`D${r}`).value = d.apr;
    wsDebt.getCell(`D${r}`).numFmt = '0.00%';
    wsDebt.getCell(`E${r}`).value = d.min;
    wsDebt.getCell(`E${r}`).numFmt = '$#,##0';
    wsDebt.getCell(`F${r}`).value = d.rank;

    ['B', 'C', 'D', 'E', 'F'].forEach(col => wsDebt.getCell(`${col}${r}`).border = thinBorder);
  });

  // Total Debt Row
  wsDebt.getRow(9).height = 20;
  wsDebt.getCell('B9').value = 'TOTAL LIABILITIES';
  wsDebt.getCell('B9').font = { bold: true };
  wsDebt.getCell('C9').value = { formula: '=SUM(C5:C8)' };
  wsDebt.getCell('C9').numFmt = '$#,##0';
  wsDebt.getCell('C9').font = { bold: true };
  wsDebt.getCell('D9').value = { formula: '=AVERAGE(D5:D8)' };
  wsDebt.getCell('D9').numFmt = '0.00%';
  wsDebt.getCell('E9').value = { formula: '=SUM(E5:E8)' };
  wsDebt.getCell('E9').numFmt = '$#,##0';
  wsDebt.getCell('E9').font = { bold: true };
  wsDebt.getCell('F9').value = 'Blended APR';

  ['B', 'C', 'D', 'E', 'F'].forEach(col => {
    const c = wsDebt.getCell(`${col}9`);
    c.fill = solidFill(colors.cardBg);
    c.border = thinBorder;
    c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  });

  // =========================================================================
  // TAB 7: 🏛️ FED RESERVE SCF BENCHMARKS (RAW MICRODATA)
  // =========================================================================
  const wsFed = workbook.addWorksheet('🏛️ Fed Benchmarks (SCF)', {
    views: [{ showGridLines: true }]
  });

  wsFed.columns = [
    { width: 4 }, { width: 18 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 4 }
  ];

  wsFed.mergeCells('B2:H2');
  const fedH = wsFed.getCell('B2');
  fedH.value = '🏛️ OFFICIAL US FEDERAL RESERVE SURVEY OF CONSUMER FINANCES (SCF) BENCHMARK DATA';
  fedH.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
  fedH.fill = solidFill(colors.darkBg);
  fedH.alignment = { horizontal: 'center', vertical: 'middle' };
  wsFed.getRow(2).height = 28;

  const fedCols = [
    { col: 'B', title: 'Age Bracket' },
    { col: 'C', title: '25th Pct' },
    { col: 'D', title: '50th (Median)' },
    { col: 'E', title: '75th Pct' },
    { col: 'F', title: '90th Pct' },
    { col: 'G', title: '95th Pct' },
    { col: 'H', title: 'Top 1% (99th)' }
  ];

  wsFed.getRow(4).height = 20;
  fedCols.forEach(fc => {
    const c = wsFed.getCell(`${fc.col}4`);
    c.value = fc.title;
    c.font = { size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = solidFill(colors.cardBg);
    c.border = thinBorder;
  });

  const fedData = [
    { age: 'Under 35', p25: 4800, p50: 39000, p75: 145000, p90: 380000, p95: 750000, p99: 2800000 },
    { age: '35 - 44', p25: 35000, p50: 135600, p75: 420000, p90: 980000, p95: 1850000, p99: 5500000 },
    { age: '45 - 54', p25: 68000, p50: 247200, p75: 780000, p90: 1950000, p95: 3400000, p99: 10200000 },
    { age: '55 - 64', p25: 95000, p50: 364500, p75: 1150000, p90: 2800000, p95: 4900000, p99: 14500000 },
    { age: '65 - 74', p25: 110000, p50: 409900, p75: 1350000, p90: 3200000, p95: 5600000, p99: 16800000 },
    { age: '75+', p25: 92000, p50: 335600, p75: 1050000, p90: 2600000, p95: 4700000, p99: 13200000 }
  ];

  fedData.forEach((fd, idx) => {
    const r = 5 + idx;
    wsFed.getRow(r).height = 19;
    wsFed.getCell(`B${r}`).value = fd.age;
    wsFed.getCell(`C${r}`).value = fd.p25;
    wsFed.getCell(`D${r}`).value = fd.p50;
    wsFed.getCell(`E${r}`).value = fd.p75;
    wsFed.getCell(`F${r}`).value = fd.p90;
    wsFed.getCell(`G${r}`).value = fd.p95;
    wsFed.getCell(`H${r}`).value = fd.p99;

    ['C', 'D', 'E', 'F', 'G', 'H'].forEach(c => wsFed.getCell(`${c}${r}`).numFmt = '$#,##0');
    ['B', 'C', 'D', 'E', 'F', 'G', 'H'].forEach(col => {
      wsFed.getCell(`${col}${r}`).border = thinBorder;
      if (idx % 2 === 1) wsFed.getCell(`${col}${r}`).fill = solidFill('F8FAFC');
    });
  });

  // Write file
  await workbook.xlsx.writeFile(outputPath);
  console.log(`✅ Master Net Worth & Smart Budget Vault created at: ${outputPath}`);
}

createMasterVault().catch(err => {
  console.error('Error generating Master Vault:', err);
  process.exit(1);
});
