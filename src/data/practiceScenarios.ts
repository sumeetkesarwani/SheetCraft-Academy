import type { AutomationSnippet, SheetPreset, SkillLevel } from '../types/excel';

export const AUTOMATION_LIBRARY: AutomationSnippet[] = [
  {
    id: 'pq-unpivot',
    title: 'Unpivot Monthly ERP Crosstab',
    language: 'Power Query M',
    description: 'Turns wide Jan–Dec columns from an ERP export into a flat Month/Amount table ready for PivotTables.',
    workflowTip: 'Select only your fixed ID columns (Cost Center, Account) and choose Unpivot Other Columns so new months added next quarter flow through automatically.',
    code: `let
  Source = Excel.CurrentWorkbook(){[Name="tblERP_Export"]}[Content],
  Typed = Table.TransformColumnTypes(Source, {{"CostCenter", type text}, {"Account", type text}}),
  Unpivoted = Table.UnpivotOtherColumns(Typed, {"CostCenter", "Account"}, "Month", "ActualAmount"),
  CleanedAmount = Table.TransformColumnTypes(Unpivoted, {{"ActualAmount", Currency.Type}}),
  FilteredZeros = Table.SelectRows(CleanedAmount, each ([ActualAmount] <> 0))
in
  FilteredZeros`,
  },
  {
    id: 'vba-consolidate',
    title: 'Consolidate Regional Tabs into Master',
    language: 'VBA Macro',
    description: 'Appends rows from every regional sheet (excluding Summary) into a single clean Master_Data sheet.',
    workflowTip: 'Always disable ScreenUpdating and Calculation before looping across 15+ tabs, then restore them in an ExitHandler.',
    code: `Sub ConsolidateRegionalTabs()
  Dim ws As Worksheet, masterWs As Worksheet
  Dim nextRow As Long, lastRow As Long

  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Set masterWs = ThisWorkbook.Sheets("Master_Data")
  masterWs.Range("A2:F10000").ClearContents
  nextRow = 2

  For Each ws In ThisWorkbook.Worksheets
    If ws.Name <> "Master_Data" And ws.Name <> "Summary" Then
      lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
      If lastRow >= 2 Then
        ws.Range("A2:F" & lastRow).Copy Destination:=masterWs.Cells(nextRow, "A")
        nextRow = nextRow + (lastRow - 1)
      End If
    End If
  Next ws

  Application.Calculation = xlCalculationAutomatic
  Application.ScreenUpdating = True
End Sub`,
  },
  {
    id: 'office-script-format',
    title: 'Format Variance Report for Board Pack',
    language: 'Office Script (TS)',
    description: 'Standardizes headers, applies currency formatting, and highlights negative variances > 5% in Excel for the Web.',
    workflowTip: 'Pair Office Scripts with Power Automate to format and email daily MIS workbooks every morning at 08:00 without opening desktop Excel.',
    code: `function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const usedRange = sheet.getUsedRange();
  if (!usedRange) return;

  const header = sheet.getRange("A1:F1");
  header.getFormat().getFill().setColor("#0F172A");
  header.getFormat().getFont().setColor("#FFFFFF");
  header.getFormat().getFont().setBold(true);

  const values = usedRange.getValues();
  for (let r = 1; r < values.length; r++) {
    const variancePct = Number(values[r][5]);
    if (variancePct < -0.05) {
      sheet.getCell(r, 5).getFormat().getFill().setColor("#FEE2E2");
      sheet.getCell(r, 5).getFormat().getFont().setColor("#991B1B");
    }
  }
  usedRange.getFormat().autofitColumns();
}`,
  },
  {
    id: 'python-outlier',
    title: 'Flag Expense Anomalies with Z-Score',
    language: 'Python (=PY)',
    description: 'Runs inside an Excel =PY() cell to pull the Expense Ledger DataFrame and isolate transactions above 2.5 standard deviations.',
    workflowTip: 'Remember that xl("tblExpenses[#All]", headers=True) evaluates in row-major order inside Excel; keep the DataFrame clean before plotting.',
    code: `df = xl("A1:D16", headers=True)
mean_amt = df["Amount"].mean()
std_amt = df["Amount"].std()

df["Z_Score"] = (df["Amount"] - mean_amt) / std_amt
outliers = df[df["Z_Score"].abs() > 2.5][["Vendor", "Department", "Amount", "Z_Score"]]
outliers.sort_values(by="Z_Score", ascending=False)`,
  },
];

export const DEFAULT_PRACTICE_PRESETS: Record<SkillLevel, SheetPreset> = {
  foundation: {
    id: 'preset-foundation',
    title: 'Tax-Inclusive Invoice Line & Threshold Lock',
    skillLevel: 'foundation',
    description: 'Lock the GST rate in $E$1 and apply a conditional volume discount in E2.',
    targetCell: 'E2',
    scenarioChips: [
      'Why did my formula break when I dragged it down?',
      'Apply 10% discount only when subtotal exceeds ₹5,000',
      'Highlight rows missing a valid customer code',
    ],
    cells: {
      A1: { raw: 'Item', format: 'header' },
      B1: { raw: 'Qty', format: 'header' },
      C1: { raw: 'Unit Price', format: 'header' },
      D1: { raw: 'Gross Total', format: 'header' },
      E1: { raw: '0.18', format: 'percent', note: 'GST Rate Lock' },
      F1: { raw: 'Net Payable', format: 'header' },

      A2: { raw: 'Ergonomic Dock' },
      B2: { raw: '4', format: 'number' },
      C2: { raw: '1450', format: 'currency' },
      D2: { raw: '=B2*C2', format: 'currency' },
      E2: { raw: '=IF(D2>5000, D2*(1-0.10)*(1+$E$1), D2*(1+$E$1))', format: 'currency' },
      F2: { raw: '=D2*(1+$E$1)', format: 'currency' },

      A3: { raw: 'USB-C Hub' },
      B3: { raw: '2', format: 'number' },
      C3: { raw: '1200', format: 'currency' },
      D3: { raw: '=B3*C3', format: 'currency' },
      E3: { raw: '=IF(D3>5000, D3*(1-0.10)*(1+$E$1), D3*(1+$E$1))', format: 'currency' },
      F3: { raw: '=D3*(1+$E$1)', format: 'currency' },

      A4: { raw: 'Monitor Arm' },
      B4: { raw: '3', format: 'number' },
      C4: { raw: '2100', format: 'currency' },
      D4: { raw: '=B4*C4', format: 'currency' },
      E4: { raw: '=IF(D4>5000, D4*(1-0.10)*(1+$E$1), D4*(1+$E$1))', format: 'currency' },
      F4: { raw: '=D4*(1+$E$1)', format: 'currency' },

      A6: { raw: 'Total Invoice', format: 'header' },
      D6: { raw: '=SUM(D2:D4)', format: 'currency' },
      E6: { raw: '=SUM(E2:E4)', format: 'currency' },
      F6: { raw: '=SUM(F2:F4)', format: 'currency' },
    },
    coachCard: {
      id: 'cc-foundation-default',
      title: 'Lock Assumptions Before Filling Down',
      steps: [
        'Keep the 18% GST rate in a single assumption cell (E1) rather than burying 0.18 inside every row.',
        'Lock both column and row with $E$1 so dragging E2 down to E4 stays anchored to the GST rate.',
        'Test the threshold live: change B2 from 4 to 2 in the grid and watch E2 drop the 10% volume discount immediately.',
      ],
      syntax: '=IF(D2>5000, D2*(1-0.10)*(1+$E$1), D2*(1+$E$1))',
      syntaxArgs: [
        { label: 'D2>5000', role: 'Checks whether gross order qualifies for the tier discount' },
        { label: '$E$1', role: 'Absolute lock on the 18% GST cell so row fills do not shift to E2, E3' },
      ],
      pitfallWarning: 'Leaving E1 unlocked turns $E$1 into E2 and E3 when copied down—producing silent 4-digit inflation across your invoice column.',
      nextStepChallenge: 'Click cell B2 and change Qty from 4 to 3, then observe how D2 and E2 recalculate.',
      suggestedFormula: '=IF(D2>5000, D2*(1-0.10)*(1+$E$1), D2*(1+$E$1))',
      targetCell: 'E2',
    },
    automationSnippets: AUTOMATION_LIBRARY.slice(0, 2),
  },
  'data-analyst': {
    id: 'preset-analyst',
    title: 'Multi-Region SKU Lookup & SUMIFS Rollup',
    skillLevel: 'data-analyst',
    description: 'Pull unit pricing cleanly with XLOOKUP and aggregate regional revenue with SUMIFS.',
    targetCell: 'G3',
    scenarioChips: [
      'Replace fragile VLOOKUP with XLOOKUP and a fallback',
      'Sum North region revenue above ₹1,500 using SUMIFS',
      'Spill a filtered list of active West orders with FILTER',
    ],
    cells: {
      A1: { raw: 'SKU', format: 'header' },
      B1: { raw: 'Region', format: 'header' },
      C1: { raw: 'Units', format: 'header' },
      D1: { raw: 'Revenue', format: 'header' },

      A2: { raw: 'SKU-101' }, B2: { raw: 'North' }, C2: { raw: '12', format: 'number' }, D2: { raw: '18000', format: 'currency' },
      A3: { raw: 'SKU-102' }, B3: { raw: 'West' }, C3: { raw: '8', format: 'number' }, D3: { raw: '14400', format: 'currency' },
      A4: { raw: 'SKU-103' }, B4: { raw: 'North' }, C4: { raw: '15', format: 'number' }, D4: { raw: '26250', format: 'currency' },
      A5: { raw: 'SKU-104' }, B5: { raw: 'South' }, C5: { raw: '5', format: 'number' }, D5: { raw: '7500', format: 'currency' },
      A6: { raw: 'SKU-105' }, B6: { raw: 'West' }, C6: { raw: '19', format: 'number' }, D6: { raw: '30400', format: 'currency' },

      F1: { raw: 'Lookup SKU', format: 'header' },
      G1: { raw: 'SKU-103' },
      F2: { raw: 'Region Found', format: 'header' },
      G2: { raw: '=XLOOKUP(G1, $A$2:$A$6, $B$2:$B$6, "Not Found")' },
      F3: { raw: 'SKU Revenue', format: 'header' },
      G3: { raw: '=XLOOKUP(G1, $A$2:$A$6, $D$2:$D$6, 0)', format: 'currency' },
      F5: { raw: 'North Total', format: 'header' },
      G5: { raw: '=SUMIFS($D$2:$D$6, $B$2:$B$6, "North")', format: 'currency' },
      F6: { raw: 'West Deals', format: 'header' },
      G6: { raw: '=COUNTIFS($B$2:$B$6, "West")', format: 'number' },
    },
    coachCard: {
      id: 'cc-analyst-default',
      title: 'Column-Safe Lookups with XLOOKUP',
      steps: [
        'Point XLOOKUP strictly at the lookup column ($A$2:$A$6) and return column ($D$2:$D$6)—never a full A:D block that breaks when someone inserts a column.',
        'Include the 4th argument (0 or "Not Found") so missing SKUs never leak #N/A into executive rollups.',
        'Change cell G1 to SKU-105 to verify G2 and G3 update instantaneously.',
      ],
      syntax: '=XLOOKUP(G1, $A$2:$A$6, $D$2:$D$6, 0)',
      syntaxArgs: [
        { label: 'G1', role: 'Search key (SKU-103)' },
        { label: '$A$2:$A$6', role: 'Locked lookup column' },
        { label: '$D$2:$D$6', role: 'Locked return column (Revenue)' },
        { label: '0', role: 'Clean fallback if SKU is absent' },
      ],
      pitfallWarning: 'VLOOKUP breaks silently if an analyst inserts a column inside the source table and forgets to increment col_index_num.',
      nextStepChallenge: 'Edit cell G1 from SKU-103 to SKU-105 and verify both G2 and G3 recalculate.',
      suggestedFormula: '=XLOOKUP(G1, $A$2:$A$6, $D$2:$D$6, 0)',
      targetCell: 'G3',
    },
    automationSnippets: AUTOMATION_LIBRARY.slice(0, 2),
  },
  'financial-modeling': {
    id: 'preset-modeling',
    title: '5-Year Cash Flow XNPV & LET Margin Check',
    skillLevel: 'financial-modeling',
    description: 'Discount irregular project cash flows with XNPV and keep multi-step ratios readable via LET.',
    targetCell: 'E2',
    scenarioChips: [
      'Calculate project XNPV at 11% WACC in E2',
      'Compute annualized XIRR return in E3',
      'Refactor nested margin math with LET variables',
    ],
    cells: {
      A1: { raw: 'Date', format: 'header' },
      B1: { raw: 'Cash Flow', format: 'header' },
      D1: { raw: 'WACC', format: 'header' },
      E1: { raw: '0.11', format: 'percent' },

      A2: { raw: '2026-01-01', format: 'date' }, B2: { raw: '-120000', format: 'currency' },
      A3: { raw: '2026-09-15', format: 'date' }, B3: { raw: '35000', format: 'currency' },
      A4: { raw: '2027-06-30', format: 'date' }, B4: { raw: '52000', format: 'currency' },
      A5: { raw: '2028-12-31', format: 'date' }, B5: { raw: '78000', format: 'currency' },

      D2: { raw: 'Project XNPV', format: 'header' },
      E2: { raw: '=XNPV(E1, B2:B5, A2:A5)', format: 'currency' },
      D3: { raw: 'Project XIRR', format: 'header' },
      E3: { raw: '=XIRR(B2:B5, A2:A5)', format: 'percent' },
      D4: { raw: 'Total Inflow', format: 'header' },
      E4: { raw: '=SUM(B3:B5)', format: 'currency' },
      D5: { raw: 'Net Cash Gain', format: 'header' },
      E5: { raw: '=LET(outflow, ABS(B2), inflow, SUM(B3:B5), (inflow-outflow)/outflow)', format: 'percent' },
    },
    coachCard: {
      id: 'cc-modeling-default',
      title: 'Use XNPV When Cash Flows Are Irregular',
      steps: [
        'Standard NPV() assumes equal 12-month intervals and discounts Period 1 by a full year—wrong for mid-year capex or uneven draws.',
        'Pass the WACC cell E1, cash flows B2:B5 (starting with the initial negative outflow), and exact dates A2:A5 into XNPV.',
        'Wrap complex return ratios inside LET() so reviewers can audit variable names instead of decoding repeated SUM() calls.',
      ],
      syntax: '=XNPV(E1, B2:B5, A2:A5)',
      syntaxArgs: [
        { label: 'E1', role: 'Annual discount rate (11.0% WACC)' },
        { label: 'B2:B5', role: 'Cash flow series starting with initial outflow' },
        { label: 'A2:A5', role: 'Exact transaction dates' },
      ],
      pitfallWarning: 'Passing a positive initial capex or omitting the Year-0 date from NPV() overstates project value by a full year of discounting.',
      nextStepChallenge: 'Raise WACC in cell E1 from 0.11 to 0.14 and watch Project XNPV in E2 compress.',
      suggestedFormula: '=XNPV(E1, B2:B5, A2:A5)',
      targetCell: 'E2',
    },
    automationSnippets: AUTOMATION_LIBRARY.slice(1, 3),
  },
  'expert-automation': {
    id: 'preset-automation',
    title: 'ERP Expense Audit & Automation Staging Sheet',
    skillLevel: 'expert-automation',
    description: 'Clean raw ERP exports with staging formulas and inspect Power Query M, VBA, Office Scripts & =PY() pipelines.',
    targetCell: 'E2',
    scenarioChips: [
      'Write a Power Query M step to unpivot monthly columns',
      'Automate workbook consolidation with VBA',
      'Flag expense outliers using Python in Excel (=PY)',
    ],
    cells: {
      A1: { raw: 'Raw Vendor', format: 'header' },
      B1: { raw: 'Department', format: 'header' },
      C1: { raw: 'Amount', format: 'header' },
      D1: { raw: 'Clean Vendor', format: 'header' },
      E1: { raw: 'Audit Flag', format: 'header' },

      A2: { raw: '  cloudstack inc ' }, B2: { raw: 'IT' }, C2: { raw: '14200', format: 'currency' },
      D2: { raw: '=PROPER(TRIM(A2))' },
      E2: { raw: '=IF(C2>AVERAGE($C$2:$C$5)*1.5, "Outlier Review", "Normal")' },

      A3: { raw: 'office  depot' }, B3: { raw: 'Admin' }, C3: { raw: '1850', format: 'currency' },
      D3: { raw: '=PROPER(TRIM(A3))' },
      E3: { raw: '=IF(C3>AVERAGE($C$2:$C$5)*1.5, "Outlier Review", "Normal")' },

      A4: { raw: '  ACME LOGISTICS ' }, B4: { raw: 'Ops' }, C4: { raw: '3200', format: 'currency' },
      D4: { raw: '=PROPER(TRIM(A4))' },
      E4: { raw: '=IF(C4>AVERAGE($C$2:$C$5)*1.5, "Outlier Review", "Normal")' },

      A5: { raw: 'metro travel' }, B5: { raw: 'Sales' }, C5: { raw: '2400', format: 'currency' },
      D5: { raw: '=PROPER(TRIM(A5))' },
      E5: { raw: '=IF(C5>AVERAGE($C$2:$C$5)*1.5, "Outlier Review", "Normal")' },

      B7: { raw: 'Mean Spend', format: 'header' },
      C7: { raw: '=AVERAGE(C2:C5)', format: 'currency' },
    },
    coachCard: {
      id: 'cc-automation-default',
      title: 'Stage Clean Types Before Automated Runs',
      steps: [
        'Use Power Query M for repeatable CSV/ERP ingestion before writing a single VBA loop or Python cell.',
        'Switch to the Automation Code tab above the grid to inspect ready-to-paste M, VBA, Office Scripts, and =PY() scripts.',
        'Keep staging columns explicit (D2 cleans whitespace via TRIM/PROPER; E2 flags variance anomalies).',
      ],
      syntax: '=IF(C2>AVERAGE($C$2:$C$5)*1.5, "Outlier Review", "Normal")',
      syntaxArgs: [
        { label: 'AVERAGE($C$2:$C$5)', role: 'Locked baseline spend mean across the ledger' },
        { label: '*1.5', role: 'Threshold multiplier before escalating to audit' },
      ],
      pitfallWarning: 'Never run VBA Selection/Activate loops on uncleaned ERP text—trailing spaces cause silent join misses in downstream lookups.',
      nextStepChallenge: 'Open the Automation Code tab above the grid and copy the Power Query M or Python snippet.',
      suggestedFormula: '=IF(C2>AVERAGE($C$2:$C$5)*1.5, "Outlier Review", "Normal")',
      targetCell: 'E2',
    },
    automationSnippets: AUTOMATION_LIBRARY,
  },
};
