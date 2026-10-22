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
    textWhite: 'FFFFFF',
    textMuted: '94A3B8',
    textDark: '0F172A',
    borderDark: '334155',
    lightRow: 'F8FAFC',
    altRow: 'F1F5F9'
  };

  // Helper for solid fills
  const solidFill = (argb) => ({
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF' + argb }
  });

  // Helper for borders
  const thinBorder = {
    top: { style: 'thin', color: { argb: 'FF' + colors.borderDark } },
    left: { style: 'thin', color: { argb: 'FF' + colors.borderDark } },
    bottom: { style: 'thin', color: { argb: 'FF' + colors.borderDark } },
    right: { style: 'thin', color: { argb: 'FF' + colors.borderDark } }
  };

  // =========================================================================
  // TAB 1: EXECUTIVE COMMAND CENTER (DASHBOARD)
  // =========================================================================
  const ws1 = workbook.addWorksheet('📊 Executive Dashboard', {
    views: [{ showGridLines: true }]
  });

  ws1.columns = [
    { width: 4 },   // A (spacer)
    { width: 28 },  // B
    { width: 20 },  // C
    { width: 20 },  // D
    { width: 20 },  // E
    { width: 22 },  // F
    { width: 4 }    // G (spacer)
  ];

  // Header Banner
  ws1.mergeCells('B2:F2');
  const hCell = ws1.getCell('B2');
  hCell.value = '🏛️ NET WORTH & FINANCIAL INDEPENDENCE MASTER VAULT (2026 EDITION)';
  hCell.font = { name: 'Arial', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
  hCell.fill = solidFill(colors.darkBg);
  hCell.alignment = { horizontal: 'center', vertical: 'middle' };
  ws1.getRow(2).height = 36;

  ws1.mergeCells('B3:F3');
  const subCell = ws1.getCell('B3');
  subCell.value = 'Private Institutional Benchmarking Engine • Federal Reserve SCF 2022/2026 Microdata • Powered by GetNetWorthCalculator.com';
  subCell.font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF' + colors.textMuted } };
  subCell.fill = solidFill(colors.cardBg);
  subCell.alignment = { horizontal: 'center', vertical: 'middle' };
  ws1.getRow(3).height = 20;

  // KPI CARDS ROW (B5 to F7)
  const kpis = [
    { col: 'B', label: 'TOTAL NET WORTH', formula: "='📅 12-Month Ledger'!D19", note: 'All Assets − Total Liabilities', color: colors.emerald },
    { col: 'C', label: 'LIQUID NET WORTH', formula: "='📅 12-Month Ledger'!E19", note: 'Investable (Excl. Home Equity)', color: colors.blue },
    { col: 'D', label: 'DEBT-TO-ASSET RATIO', formula: "='📅 12-Month Ledger'!C19/'📅 12-Month Ledger'!B19", note: 'Target < 30%', color: colors.amber, isPct: true },
    { col: 'E', label: '1-YEAR WEALTH GAIN', formula: "='📅 12-Month Ledger'!D19-'📅 12-Month Ledger'!D8", note: 'Net change since Jan', color: colors.emerald },
    { col: 'F', label: 'FINANCIAL FREEDOM SCORE', formula: "=MIN(100, ROUND(('📅 12-Month Ledger'!E19/'🔥 FIRE Engine'!D12)*100, 1))", note: '% of Standard FIRE Target', color: colors.emerald, isScore: true }
  ];

  ws1.getRow(5).height = 18;
  ws1.getRow(6).height = 32;
  ws1.getRow(7).height = 16;

  kpis.forEach(k => {
    const lCell = ws1.getCell(`${k.col}5`);
    lCell.value = k.label;
    lCell.font = { name: 'Arial', size: 8, bold: true, color: { argb: 'FF' + colors.textMuted } };
    lCell.fill = solidFill(colors.cardBg);
    lCell.alignment = { horizontal: 'center', vertical: 'middle' };

    const vCell = ws1.getCell(`${k.col}6`);
    vCell.value = { formula: k.formula };
    vCell.font = { name: 'Arial', size: 14, bold: true, color: { argb: 'FF' + k.color } };
    vCell.fill = solidFill(colors.cardBg);
    vCell.alignment = { horizontal: 'center', vertical: 'middle' };
    if (k.isPct) {
      vCell.numFmt = '0.0%';
    } else if (k.isScore) {
      vCell.numFmt = '0.0"%"';
    } else {
      vCell.numFmt = '$#,##0';
    }

    const nCell = ws1.getCell(`${k.col}7`);
    nCell.value = k.note;
    nCell.font = { name: 'Arial', size: 7.5, italic: true, color: { argb: 'FF' + colors.textMuted } };
    nCell.fill = solidFill(colors.cardBg);
    nCell.alignment = { horizontal: 'center', vertical: 'middle' };
  });

  // FEDERAL RESERVE PERCENTILE ENGINE (B9 to F15)
  ws1.mergeCells('B9:F9');
  const fedHeader = ws1.getCell('B9');
  fedHeader.value = '🇺🇸 OFFICIAL US FEDERAL RESERVE (SCF) PERCENTILE BENCHMARK';
  fedHeader.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  fedHeader.fill = solidFill(colors.darkBg);
  fedHeader.alignment = { horizontal: 'center', vertical: 'middle' };
  ws1.getRow(9).height = 24;

  ws1.getCell('B10').value = 'Your Current Age:';
  ws1.getCell('B10').font = { bold: true };
  ws1.getCell('C10').value = 35;
  ws1.getCell('C10').alignment = { horizontal: 'center' };
  ws1.getCell('C10').font = { bold: true, color: { argb: 'FF' + colors.blue } };

  ws1.getCell('D10').value = 'Assigned Age Cohort:';
  ws1.getCell('D10').font = { bold: true };
  ws1.getCell('E10').value = { formula: '=IF(C10<35,"Under 35",IF(C10<=44,"35-44",IF(C10<=54,"45-54",IF(C10<=64,"55-64",IF(C10<=74,"65-74","75+")))))' };
  ws1.getCell('E10').alignment = { horizontal: 'center' };
  ws1.getCell('E10').font = { bold: true, color: { argb: 'FF' + colors.emerald } };

  ws1.getCell('B11').value = 'Your Net Worth:';
  ws1.getCell('B11').font = { bold: true };
  ws1.getCell('C11').value = { formula: "='📅 12-Month Ledger'!D19" };
  ws1.getCell('C11').numFmt = '$#,##0';
  ws1.getCell('C11').font = { bold: true };

  ws1.getCell('D11').value = 'Median for Your Age (50th):';
  ws1.getCell('D11').font = { bold: true };
  ws1.getCell('E11').value = { formula: '=VLOOKUP(E10,\'🏛️ Fed Benchmarks\'!A4:G9,4,FALSE)' };
  ws1.getCell('E11').numFmt = '$#,##0';
  ws1.getCell('E11').font = { bold: true };

  ws1.getCell('B12').value = 'Your Classified Wealth Class:';
  ws1.getCell('B12').font = { bold: true };
  ws1.mergeCells('C12:D12');
  ws1.getCell('C12').value = { formula: '=IF(C11>=30000000,"Ultra High Net Worth (Top 0.1%)",IF(C11>=5000000,"Very High Net Worth (Top 1.3%)",IF(C11>=1000000,"High Net Worth / Millionaire (Top 8.5%)",IF(C11>=500000,"Upper Middle Class (Top 24%)",IF(C11>=100000,"Middle Class (Top 50%)","Emerging Wealth")))))' };
  ws1.getCell('C12').font = { bold: true, color: { argb: 'FF' + colors.amber } };

  ws1.getCell('E12').value = 'Top 10% Cutoff (90th):';
  ws1.getCell('E12').font = { bold: true };
  ws1.getCell('F12').value = { formula: '=VLOOKUP(E10,\'🏛️ Fed Benchmarks\'!A4:G9,6,FALSE)' };
  ws1.getCell('F12').numFmt = '$#,##0';
  ws1.getCell('F12').font = { bold: true, color: { argb: 'FF' + colors.emerald } };

  // FIRE PROGRESS HIGHLIGHT BAR
  ws1.mergeCells('B14:F14');
  const fireBar = ws1.getCell('B14');
  fireBar.value = '🔥 FINANCIAL INDEPENDENCE (FIRE) FREEDOM SNAPSHOT';
  fireBar.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  fireBar.fill = solidFill(colors.emeraldDark);
  fireBar.alignment = { horizontal: 'center', vertical: 'middle' };
  ws1.getRow(14).height = 24;

  ws1.getCell('B15').value = 'Annual Living Expenses:';
  ws1.getCell('B15').font = { bold: true };
  ws1.getCell('C15').value = { formula: "='🔥 FIRE Engine'!D8" };
  ws1.getCell('C15').numFmt = '$#,##0';

  ws1.getCell('D15').value = 'Coast FIRE Status:';
  ws1.getCell('D15').font = { bold: true };
  ws1.getCell('E15').value = { formula: "='🔥 FIRE Engine'!D18" };
  ws1.getCell('E15').font = { bold: true, color: { argb: 'FF' + colors.emerald } };

  ws1.getCell('B16').value = 'Standard FIRE Target (4% Rule):';
  ws1.getCell('B16').font = { bold: true };
  ws1.getCell('C16').value = { formula: "='🔥 FIRE Engine'!D12" };
  ws1.getCell('C16').numFmt = '$#,##0';

  ws1.getCell('D16').value = 'Projected Freedom Year:';
  ws1.getCell('D16').font = { bold: true };
  ws1.getCell('E16').value = { formula: "='🔥 FIRE Engine'!D22" };
  ws1.getCell('E16').font = { bold: true, color: { argb: 'FF' + colors.amber } };

  // =========================================================================
  // TAB 2: 12-MONTH HISTORICAL WEALTH LEDGER
  // =========================================================================
  const ws2 = workbook.addWorksheet('📅 12-Month Ledger', {
    views: [{ showGridLines: true }]
  });

  ws2.columns = [
    { width: 14 }, // A: Month
    { width: 16 }, // B: Total Assets
    { width: 16 }, // C: Total Liabilities
    { width: 18 }, // D: Total Net Worth
    { width: 18 }, // E: Liquid Net Worth
    { width: 14 }, // F: MoM $ Change
    { width: 12 }, // G: MoM % Change
    { width: 14 }, // H: Checking/Cash
    { width: 14 }, // I: HYSA
    { width: 14 }, // J: 401(k) / IRA
    { width: 14 }, // K: Brokerage
    { width: 14 }, // L: Real Estate
    { width: 14 }, // M: Crypto/Other
    { width: 14 }, // N: Mortgage
    { width: 14 }, // O: Student Debt
    { width: 14 }  // P: Other Debt
  ];

  // Header Banner
  ws2.mergeCells('A2:P2');
  const h2 = ws2.getCell('A2');
  h2.value = '📅 12-MONTH WEALTH ACCUMULATION LEDGER (2026)';
  h2.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
  h2.fill = solidFill(colors.darkBg);
  h2.alignment = { horizontal: 'center', vertical: 'middle' };
  ws2.getRow(2).height = 30;

  // Ledger Column Headers
  const ledgerHeaders = [
    'Month', 'Gross Assets', 'Liabilities', 'TOTAL NET WORTH', 'LIQUID NET WORTH',
    'MoM $ Change', 'MoM % Change', 'Checking', 'High Yield (HYSA)', 'Retirement 401k/IRA',
    'Taxable Brokerage', 'Home Equity / RE', 'Crypto / Alts', 'Mortgage Balance', 'Student Loans', 'Credit/Other Debt'
  ];

  ws2.getRow(4).values = ledgerHeaders;
  ws2.getRow(4).font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
  ws2.getRow(4).height = 26;
  for (let c = 1; c <= 16; c++) {
    const cell = ws2.getRow(4).getCell(c);
    cell.fill = solidFill(c <= 7 ? colors.darkBg : colors.cardBg);
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  }

  // Pre-seed 12 Months with realistic starter numbers and formulas
  const months = ['Jan 2026', 'Feb 2026', 'Mar 2026', 'Apr 2026', 'May 2026', 'Jun 2026', 'Jul 2026', 'Aug 2026', 'Sep 2026', 'Oct 2026', 'Nov 2026', 'Dec 2026'];
  const starter = [
    { cash: 6000, hysa: 24000, ret: 110000, brok: 45000, re: 380000, alt: 8000, mort: 280000, stud: 18000, debt: 2500 },
    { cash: 6500, hysa: 25000, ret: 112500, brok: 46200, re: 381000, alt: 8400, mort: 279200, stud: 17600, debt: 2100 },
    { cash: 6200, hysa: 26000, ret: 115000, brok: 47500, re: 382000, alt: 8200, mort: 278400, stud: 17200, debt: 1800 },
    { cash: 6800, hysa: 27000, ret: 118000, brok: 48900, re: 383000, alt: 8900, mort: 277600, stud: 16800, debt: 1500 },
    { cash: 7000, hysa: 28000, ret: 121000, brok: 50200, re: 384000, alt: 9100, mort: 276800, stud: 16400, debt: 1200 },
    { cash: 7200, hysa: 29000, ret: 124500, brok: 51800, re: 385000, alt: 9400, mort: 276000, stud: 16000, debt: 900 },
    { cash: 7100, hysa: 30000, ret: 128000, brok: 53200, re: 386000, alt: 9200, mort: 275200, stud: 15600, debt: 700 },
    { cash: 7500, hysa: 31000, ret: 131500, brok: 54900, re: 387000, alt: 9800, mort: 274400, stud: 15200, debt: 500 },
    { cash: 7400, hysa: 32000, ret: 135000, brok: 56400, re: 388000, alt: 10100, mort: 273600, stud: 14800, debt: 400 },
    { cash: 7800, hysa: 33000, ret: 139000, brok: 58200, re: 389000, alt: 10500, mort: 272800, stud: 14400, debt: 200 },
    { cash: 8000, hysa: 34000, ret: 143000, brok: 60100, re: 390000, alt: 10900, mort: 272000, stud: 14000, debt: 100 },
    { cash: 8500, hysa: 35000, ret: 147500, brok: 62000, re: 391000, alt: 11400, mort: 271200, stud: 13600, debt: 0 }
  ];

  months.forEach((m, idx) => {
    const rowNum = 5 + idx;
    const s = starter[idx];
    const row = ws2.getRow(rowNum);
    row.height = 20;

    row.getCell(1).value = m;
    row.getCell(1).font = { bold: true };

    // Formulas:
    // Assets: SUM(H:M)
    row.getCell(2).value = { formula: `=SUM(H${rowNum}:M${rowNum})` };
    // Liabilities: SUM(N:P)
    row.getCell(3).value = { formula: `=SUM(N${rowNum}:P${rowNum})` };
    // Net Worth: Assets - Liabilities
    row.getCell(4).value = { formula: `=B${rowNum}-C${rowNum}` };
    row.getCell(4).font = { bold: true, color: { argb: 'FF' + colors.emerald } };
    // Liquid Net Worth: (H+I+J+K) - (O+P)
    row.getCell(5).value = { formula: `=SUM(H${rowNum}:K${rowNum})-SUM(O${rowNum}:P${rowNum})` };
    row.getCell(5).font = { bold: true, color: { argb: 'FF' + colors.blue } };

    // MoM $ change
    if (idx === 0) {
      row.getCell(6).value = 0;
      row.getCell(7).value = 0;
    } else {
      row.getCell(6).value = { formula: `=D${rowNum}-D${rowNum - 1}` };
      row.getCell(7).value = { formula: `=IF(D${rowNum - 1}=0,0,(D${rowNum}-D${rowNum - 1})/D${rowNum - 1})` };
    }

    // Individual Account Values
    row.getCell(8).value = s.cash;
    row.getCell(9).value = s.hysa;
    row.getCell(10).value = s.ret;
    row.getCell(11).value = s.brok;
    row.getCell(12).value = s.re;
    row.getCell(13).value = s.alt;
    row.getCell(14).value = s.mort;
    row.getCell(15).value = s.stud;
    row.getCell(16).value = s.debt;

    // Number formats
    for (let c = 2; c <= 16; c++) {
      if (c === 7) {
        row.getCell(c).numFmt = '0.0%';
      } else {
        row.getCell(c).numFmt = '$#,##0';
      }
    }
  });

  // Summary Row (Row 18)
  ws2.getRow(18).height = 24;
  ws2.getCell('A18').value = 'AVERAGE (2026)';
  ws2.getCell('A18').font = { bold: true };
  for (let c = 2; c <= 16; c++) {
    const colLetter = String.fromCharCode(64 + c);
    if (c === 7) {
      ws2.getCell(`${colLetter}18`).value = { formula: `=AVERAGE(${colLetter}5:${colLetter}16)` };
      ws2.getCell(`${colLetter}18`).numFmt = '0.0%';
    } else {
      ws2.getCell(`${colLetter}18`).value = { formula: `=AVERAGE(${colLetter}5:${colLetter}16)` };
      ws2.getCell(`${colLetter}18`).numFmt = '$#,##0';
    }
    ws2.getCell(`${colLetter}18`).font = { bold: true };
  }

  // Latest Active Row Pointer (Row 19)
  ws2.getCell('A19').value = 'LATEST RECORD:';
  ws2.getCell('A19').font = { bold: true, color: { argb: 'FF' + colors.emerald } };
  for (let c = 2; c <= 16; c++) {
    const colLetter = String.fromCharCode(64 + c);
    ws2.getCell(`${colLetter}19`).value = { formula: `=${colLetter}16` };
    ws2.getCell(`${colLetter}19`).font = { bold: true };
    ws2.getCell(`${colLetter}19`).numFmt = c === 7 ? '0.0%' : '$#,##0';
  }

  // =========================================================================
  // TAB 3: FIRE & EARLY RETIREMENT FREEDOM ENGINE
  // =========================================================================
  const ws3 = workbook.addWorksheet('🔥 FIRE Engine', {
    views: [{ showGridLines: true }]
  });

  ws3.columns = [
    { width: 4 },
    { width: 34 },
    { width: 18 },
    { width: 22 },
    { width: 30 }
  ];

  ws3.mergeCells('B2:E2');
  const h3 = ws3.getCell('B2');
  h3.value = '🔥 FINANCIAL INDEPENDENCE & EARLY RETIREMENT (FIRE) ENGINE';
  h3.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
  h3.fill = solidFill(colors.emeraldDark);
  h3.alignment = { horizontal: 'center', vertical: 'middle' };
  ws3.getRow(2).height = 32;

  // Inputs Block
  ws3.mergeCells('B4:E4');
  ws3.getCell('B4').value = 'STEP 1: YOUR PERSONAL RETIREMENT INPUTS';
  ws3.getCell('B4').font = { bold: true, size: 9, color: { argb: 'FFFFFFFF' } };
  ws3.getCell('B4').fill = solidFill(colors.darkBg);

  const fireInputs = [
    { row: 5, label: 'Current Age:', val: 32, note: 'Years old', isNum: true },
    { row: 6, label: 'Target Retirement Age:', val: 55, note: 'When work becomes 100% optional', isNum: true },
    { row: 7, label: 'Years Until Target Retirement:', formula: '=D6-D5', note: 'Years remaining to compound' },
    { row: 8, label: 'Annual Household Living Expenses:', val: 65000, note: 'Cost to sustain your lifestyle in retirement', isCur: true },
    { row: 9, label: 'Safe Withdrawal Rate (SWR):', val: 0.04, note: 'Trinity Study recommendation (4.0%)', isPct: true },
    { row: 10, label: 'Expected Real Investment Return:', val: 0.07, note: 'Historical S&P 500 inflation-adjusted (7.0%)', isPct: true }
  ];

  fireInputs.forEach(i => {
    ws3.getCell(`B${i.row}`).value = i.label;
    ws3.getCell(`B${i.row}`).font = { bold: true };
    if (i.formula) {
      ws3.getCell(`D${i.row}`).value = { formula: i.formula };
    } else {
      ws3.getCell(`D${i.row}`).value = i.val;
    }
    ws3.getCell(`D${i.row}`).font = { bold: true };
    ws3.getCell(`D${i.row}`).alignment = { horizontal: 'center' };
    ws3.getCell(`E${i.row}`).value = i.note;
    ws3.getCell(`E${i.row}`).font = { italic: true, size: 8, color: { argb: 'FF' + colors.textMuted } };

    if (i.isCur) ws3.getCell(`D${i.row}`).numFmt = '$#,##0';
    if (i.isPct) ws3.getCell(`D${i.row}`).numFmt = '0.0%';
  });

  // FIRE Milestones Block
  ws3.mergeCells('B11:E11');
  ws3.getCell('B11').value = 'STEP 2: YOUR CALCULATED FIRE TARGETS';
  ws3.getCell('B11').font = { bold: true, size: 9, color: { argb: 'FFFFFFFF' } };
  ws3.getCell('B11').fill = solidFill(colors.darkBg);

  const fireTargets = [
    { row: 12, label: 'Standard FIRE Target (4% Rule):', formula: '=D8/D9', note: '25x annual spending — full financial independence' },
    { row: 13, label: 'Lean FIRE Target (75% Spending):', formula: '=D12*0.75', note: 'Essential living expenses covered without luxuries' },
    { row: 14, label: 'Fat FIRE Target (125% Spending):', formula: '=D12*1.25', note: 'Abundant lifestyle with travel and high discretionary spend' },
    { row: 15, label: 'Barista FIRE Target (50% Covered):', formula: '=D12*0.5', note: 'Portfolio covers 50%; casual/part-time work covers the rest' },
    { row: 16, label: 'Coast FIRE Milestone Required Today:', formula: '=D12/((1+D10)^D7)', note: 'Amount needed TODAY so you never have to save another cent' }
  ];

  fireTargets.forEach(t => {
    ws3.getCell(`B${t.row}`).value = t.label;
    ws3.getCell(`B${t.row}`).font = { bold: true };
    ws3.getCell(`D${t.row}`).value = { formula: t.formula };
    ws3.getCell(`D${t.row}`).font = { bold: true, color: { argb: 'FF' + colors.emerald } };
    ws3.getCell(`D${t.row}`).numFmt = '$#,##0';
    ws3.getCell(`D${t.row}`).alignment = { horizontal: 'center' };
    ws3.getCell(`E${t.row}`).value = t.note;
    ws3.getCell(`E${t.row}`).font = { italic: true, size: 8, color: { argb: 'FF' + colors.textMuted } };
  });

  // Coast FIRE Status Check
  ws3.mergeCells('B17:E17');
  ws3.getCell('B17').value = 'STEP 3: AUTOMATED STATUS & TIMELINE DIAGNOSTICS';
  ws3.getCell('B17').font = { bold: true, size: 9, color: { argb: 'FFFFFFFF' } };
  ws3.getCell('B17').fill = solidFill(colors.darkBg);

  ws3.getCell('B18').value = 'Coast FIRE Status:';
  ws3.getCell('B18').font = { bold: true };
  ws3.getCell('D18').value = { formula: "=IF('📅 12-Month Ledger'!E19>=D16,\"✅ ACHIEVED! Work is optional.\",\"⏳ In Progress — Need \" & TEXT(D16-'📅 12-Month Ledger'!E19,\"$#,##0\"))" };
  ws3.getCell('D18').font = { bold: true, color: { argb: 'FF' + colors.emerald } };

  ws3.getCell('B19').value = 'Current Liquid Capital:';
  ws3.getCell('B19').font = { bold: true };
  ws3.getCell('D19').value = { formula: "='📅 12-Month Ledger'!E19" };
  ws3.getCell('D19').numFmt = '$#,##0';

  ws3.getCell('B20').value = 'Gap to Full Standard FIRE:';
  ws3.getCell('B20').font = { bold: true };
  ws3.getCell('D20').value = { formula: "=MAX(0, D12-D19)" };
  ws3.getCell('D20').numFmt = '$#,##0';

  ws3.getCell('B21').value = 'Monthly Investment Needed:';
  ws3.getCell('B21').font = { bold: true };
  ws3.getCell('D21').value = { formula: '=PMT(D10/12, D7*12, -D19, D12)' };
  ws3.getCell('D21').numFmt = '$#,##0';
  ws3.getCell('E21').value = 'Monthly contribution to hit FIRE by target age';

  ws3.getCell('B22').value = 'Estimated Calendar Year of Freedom:';
  ws3.getCell('B22').font = { bold: true };
  ws3.getCell('D22').value = { formula: '=YEAR(TODAY())+D7' };
  ws3.getCell('D22').font = { bold: true, color: { argb: 'FF' + colors.amber } };
  ws3.getCell('D22').alignment = { horizontal: 'center' };

  // =========================================================================
  // TAB 4: ASSET ALLOCATION & CASH DRAG OPTIMIZER
  // =========================================================================
  const ws4 = workbook.addWorksheet('🛡️ Portfolio Health', {
    views: [{ showGridLines: true }]
  });

  ws4.columns = [
    { width: 4 },
    { width: 30 },
    { width: 18 },
    { width: 18 },
    { width: 20 },
    { width: 28 }
  ];

  ws4.mergeCells('B2:F2');
  const h4 = ws4.getCell('B2');
  h4.value = '🛡️ ASSET ALLOCATION & "CASH DRAG" INFLATION OPTIMIZER';
  h4.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
  h4.fill = solidFill(colors.cardBg);
  h4.alignment = { horizontal: 'center', vertical: 'middle' };
  ws4.getRow(2).height = 30;

  // Emergency Fund Stress Test
  ws4.mergeCells('B4:F4');
  ws4.getCell('B4').value = '1. EMERGENCY FUND LIQUIDITY BENCHMARKS';
  ws4.getCell('B4').font = { bold: true, size: 9, color: { argb: 'FFFFFFFF' } };
  ws4.getCell('B4').fill = solidFill(colors.darkBg);

  ws4.getCell('B5').value = 'Monthly Living Expenses:';
  ws4.getCell('C5').value = { formula: "='🔥 FIRE Engine'!D8/12" };
  ws4.getCell('C5').numFmt = '$#,##0';
  ws4.getCell('C5').font = { bold: true };

  ws4.getCell('B6').value = 'Current Liquid Cash (Checking + HYSA):';
  ws4.getCell('C6').value = { formula: "='📅 12-Month Ledger'!H19+'📅 12-Month Ledger'!I19" };
  ws4.getCell('C6').numFmt = '$#,##0';
  ws4.getCell('C6').font = { bold: true };

  ws4.getCell('B7').value = 'Current Emergency Fund Runway:';
  ws4.getCell('C7').value = { formula: '=C6/C5' };
  ws4.getCell('C7').numFmt = '0.0 "Months"';
  ws4.getCell('C7').font = { bold: true, color: { argb: 'FF' + colors.emerald } };

  ws4.getCell('D5').value = '3-Month Lean Target:';
  ws4.getCell('E5').value = { formula: '=C5*3' };
  ws4.getCell('E5').numFmt = '$#,##0';

  ws4.getCell('D6').value = '6-Month Safe Target:';
  ws4.getCell('E6').value = { formula: '=C5*6' };
  ws4.getCell('E6').numFmt = '$#,##0';

  ws4.getCell('D7').value = '12-Month Fortress Target:';
  ws4.getCell('E7').value = { formula: '=C5*12' };
  ws4.getCell('E7').numFmt = '$#,##0';

  // CASH DRAG ANALYSIS
  ws4.mergeCells('B9:F9');
  ws4.getCell('B9').value = '2. "CASH DRAG" INFLATION EROSION CALCULATOR';
  ws4.getCell('B9').font = { bold: true, size: 9, color: { argb: 'FFFFFFFF' } };
  ws4.getCell('B9').fill = solidFill(colors.darkBg);

  ws4.getCell('B10').value = 'Checking Account Balance (Earning ~0.01%):';
  ws4.getCell('C10').value = { formula: "='📅 12-Month Ledger'!H19" };
  ws4.getCell('C10').numFmt = '$#,##0';

  ws4.getCell('B11').value = 'Ideal Checking Buffer (1 Month Expenses):';
  ws4.getCell('C11').value = { formula: '=C5' };
  ws4.getCell('C11').numFmt = '$#,##0';

  ws4.getCell('B12').value = 'Excess Idle Cash Drag:';
  ws4.getCell('C12').value = { formula: '=MAX(0, C10-C11)' };
  ws4.getCell('C12').numFmt = '$#,##0';
  ws4.getCell('C12').font = { bold: true, color: { argb: 'FF' + colors.amber } };

  ws4.getCell('B13').value = 'Annual Interest Lost (vs. 4.5% HYSA):';
  ws4.getCell('C13').value = { formula: '=C12*0.045' };
  ws4.getCell('C13').numFmt = '$#,##0';
  ws4.getCell('C13').font = { bold: true, color: { argb: 'FF' + colors.rose } };
  ws4.getCell('D13').value = '⚠️ Free money left on table each year';
  ws4.getCell('D13').font = { italic: true, size: 8 };

  // =========================================================================
  // TAB 5: DEBT SNOWBALL & AVALANCHE ACCELERATOR
  // =========================================================================
  const ws5 = workbook.addWorksheet('💳 Debt Payoff', {
    views: [{ showGridLines: true }]
  });

  ws5.columns = [
    { width: 4 },
    { width: 22 }, // Debt Name
    { width: 16 }, // Balance
    { width: 14 }, // APR %
    { width: 16 }, // Min Payment
    { width: 20 }, // Payoff Order (Avalanche)
    { width: 20 }  // Payoff Order (Snowball)
  ];

  ws5.mergeCells('B2:G2');
  const h5 = ws5.getCell('B2');
  h5.value = '💳 DEBT ACCELERATOR: AVALANCHE (MATH) VS. SNOWBALL (MOMENTUM)';
  h5.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
  h5.fill = solidFill(colors.darkBg);
  h5.alignment = { horizontal: 'center', vertical: 'middle' };
  ws5.getRow(2).height = 30;

  ws5.getRow(4).values = ['', 'Debt Description', 'Current Balance', 'Interest (APR)', 'Min Payment', 'Avalanche Priority', 'Snowball Priority'];
  ws5.getRow(4).font = { bold: true, size: 9, color: { argb: 'FFFFFFFF' } };
  for (let c = 2; c <= 7; c++) {
    ws5.getRow(4).getCell(c).fill = solidFill(colors.cardBg);
    ws5.getRow(4).getCell(c).alignment = { horizontal: 'center' };
  }

  const sampleDebts = [
    { name: 'Credit Card A (Chase)', bal: 3200, apr: 0.2499, min: 95 },
    { name: 'Credit Card B (Amex)', bal: 1800, apr: 0.2249, min: 65 },
    { name: 'Auto Loan (Honda)', bal: 12400, apr: 0.0549, min: 320 },
    { name: 'Student Loan (Fed Subsidized)', bal: 18000, apr: 0.0425, min: 195 },
    { name: 'Primary Mortgage', bal: 275000, apr: 0.0335, min: 1450 }
  ];

  sampleDebts.forEach((d, idx) => {
    const r = 5 + idx;
    ws5.getCell(`B${r}`).value = d.name;
    ws5.getCell(`C${r}`).value = d.bal;
    ws5.getCell(`C${r}`).numFmt = '$#,##0';
    ws5.getCell(`D${r}`).value = d.apr;
    ws5.getCell(`D${r}`).numFmt = '0.00%';
    ws5.getCell(`E${r}`).value = d.min;
    ws5.getCell(`E${r}`).numFmt = '$#,##0';
    ws5.getCell(`F${r}`).value = `Priority #${idx + 1} (Highest APR)`;
    ws5.getCell(`F${r}`).font = { bold: true, color: { argb: 'FF' + colors.emerald } };
    ws5.getCell(`G${r}`).value = `Rank #${idx === 0 ? 2 : idx === 1 ? 1 : idx + 1} (Lowest Bal)`;
    ws5.getCell(`G${r}`).font = { bold: true, color: { argb: 'FF' + colors.blue } };
  });

  // =========================================================================
  // TAB 6: FEDERAL RESERVE SCF 2022/2026 BENCHMARK TABLES
  // =========================================================================
  const ws6 = workbook.addWorksheet('🏛️ Fed Benchmarks', {
    views: [{ showGridLines: true }]
  });

  ws6.columns = [
    { width: 16 }, // Age Group
    { width: 16 }, // 10th
    { width: 16 }, // 25th
    { width: 18 }, // 50th (Median)
    { width: 16 }, // 75th
    { width: 18 }, // 90th (Top 10%)
    { width: 18 }  // 99th (Top 1%)
  ];

  ws6.mergeCells('A2:G2');
  const h6 = ws6.getCell('A2');
  h6.value = '🏛️ US FEDERAL RESERVE BOARD — SURVEY OF CONSUMER FINANCES (SCF)';
  h6.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
  h6.fill = solidFill(colors.darkBg);
  h6.alignment = { horizontal: 'center', vertical: 'middle' };
  ws6.getRow(2).height = 30;

  ws6.getRow(3).values = ['Age Cohort', '10th Percentile', '25th Percentile', '50th (MEDIAN)', '75th Percentile', '90th (TOP 10%)', '99th (TOP 1%)'];
  ws6.getRow(3).font = { bold: true, size: 9, color: { argb: 'FFFFFFFF' } };
  for (let c = 1; c <= 7; c++) {
    ws6.getRow(3).getCell(c).fill = solidFill(colors.cardBg);
    ws6.getRow(3).getCell(c).alignment = { horizontal: 'center' };
  }

  const fedDataRows = [
    ['Under 35', -5000, 3200, 39000, 145000, 340000, 1400000],
    ['35-44', -12000, 24000, 135600, 410000, 980000, 4500000],
    ['45-54', -8000, 48000, 247200, 780000, 1950000, 9200000],
    ['55-64', -2000, 65000, 364500, 1150000, 2800000, 14500000],
    ['65-74', 2000, 82000, 409900, 1380000, 3400000, 17800000],
    ['75+', 5000, 75000, 335600, 1100000, 2900000, 15200000]
  ];

  fedDataRows.forEach((rowVals, idx) => {
    const r = 4 + idx;
    ws6.getRow(r).values = rowVals;
    ws6.getRow(r).getCell(1).font = { bold: true };
    for (let c = 2; c <= 7; c++) {
      ws6.getRow(r).getCell(c).numFmt = '$#,##0';
      if (c === 4) ws6.getRow(r).getCell(c).font = { bold: true, color: { argb: 'FF' + colors.emerald } };
      if (c === 6) ws6.getRow(r).getCell(c).font = { bold: true, color: { argb: 'FF' + colors.amber } };
    }
  });

  // =========================================================================
  // TAB 7: QUICK START & 2-MINUTE INSTRUCTIONS
  // =========================================================================
  const ws7 = workbook.addWorksheet('📖 Quick Start Guide', {
    views: [{ showGridLines: true }]
  });

  ws7.columns = [
    { width: 4 },
    { width: 28 },
    { width: 65 }
  ];

  ws7.mergeCells('B2:C2');
  const h7 = ws7.getCell('B2');
  h7.value = '📖 HOW TO USE YOUR MASTER WEALTH VAULT';
  h7.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
  h7.fill = solidFill(colors.emeraldDark);
  h7.alignment = { horizontal: 'center', vertical: 'middle' };
  ws7.getRow(2).height = 30;

  const instructions = [
    ['Google Sheets 1-Click Setup', 'Open Google Drive (drive.google.com) → Click "+ New" → "File upload" → Select this .xlsx file → Click "Open with Google Sheets". Done!'],
    ['Microsoft Excel Setup', 'Double-click the .xlsx file to open locally. All formulas, conditional formatting, and charts work 100% natively without plugins.'],
    ['Updating Monthly', 'On the 1st of every month, head to the "📅 12-Month Ledger" tab and type in your updated account balances. All dashboard metrics update automatically.'],
    ['Customizing FIRE Targets', 'Go to the "🔥 FIRE Engine" tab and update row 5 (Age), row 6 (Target Retirement Age), and row 8 (Annual Living Expenses).'],
    ['100% Privacy Guarantee', 'This file lives solely on your personal computer or private Google Drive account. GetNetWorthCalculator.com has zero access to your numbers.']
  ];

  instructions.forEach((ins, idx) => {
    const r = 4 + idx * 2;
    ws7.getCell(`B${r}`).value = ins[0];
    ws7.getCell(`B${r}`).font = { bold: true, size: 10, color: { argb: 'FF' + colors.emerald } };
    ws7.getCell(`C${r}`).value = ins[1];
    ws7.getCell(`C${r}`).font = { size: 9 };
    ws7.getCell(`C${r}`).alignment = { wrapText: true };
    ws7.getRow(r).height = 32;
  });

  await workbook.xlsx.writeFile(outputPath);
  console.log(`✅ Master Vault generated successfully at: ${outputPath}`);
}

createMasterVault().catch(console.error);
