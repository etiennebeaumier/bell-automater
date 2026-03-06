/**
 * BCECN Master File – Office Script
 *
 * Called by Power Automate after the Azure Function parses a PDF.
 * Appends a new pricing row to the "Pricing" sheet and rebuilds the
 * four yield-curve charts in "Summary Charts".
 *
 * Power Automate "Run script" action parameters:
 *   workbook  – bound automatically to the Master File on SharePoint/OneDrive
 *   dataJson  – string: the JSON body returned by the Azure Function
 *               e.g. '{"date":"2026-03-02","bank":"TD","cad_spread_3y":85,...}'
 *
 * Returns a status string that Power Automate can log or branch on.
 */

// ---------------------------------------------------------------------------
// Column map: 1-indexed column number → JSON field name
// Mirrors COLUMN_MAP in excel_writer.py exactly.
// ---------------------------------------------------------------------------
const COLUMN_MAP: { [col: number]: string } = {
  3:  "cad_spread_3y",   4:  "cad_yield_3y",
  5:  "cad_spread_5y",   6:  "cad_yield_5y",
  7:  "cad_spread_7y",   8:  "cad_yield_7y",
  9:  "cad_spread_10y",  10: "cad_yield_10y",
  11: "cad_spread_30y",  12: "cad_yield_30y",
  13: "usd_spread_3y",   14: "usd_yield_3y",
  15: "usd_spread_5y",   16: "usd_yield_5y",
  17: "usd_spread_7y",   18: "usd_yield_7y",
  19: "usd_spread_10y",  20: "usd_yield_10y",
  21: "usd_spread_30y",  22: "usd_yield_30y",
  23: "cad_nc5_spread",  24: "cad_nc5_coupon",
  25: "cad_nc10_spread", 26: "cad_nc10_coupon",
  27: "usd_nc5_spread",  28: "usd_nc5_coupon",
  29: "usd_nc10_spread", 30: "usd_nc10_coupon",
};

// Columns that hold yield/coupon percentages (stored as decimals, shown as %).
// Mirrors the pct_columns set in excel_writer.py.
const PCT_COLUMNS = new Set([4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30]);

// Staging table anchor rows in Summary Charts (same as excel_writer.py).
// Charts are positioned using top-left cell references.
const CHART_CONFIGS = [
  {
    title:         "Bell Canada - CAD New Issue Spread Curve (bps)",
    cols:          [3, 5, 7, 9, 11],
    yLabel:        "Spread (bps)",
    isPct:         false,
    chartAnchorRow: 1, chartAnchorCol: 1,   // A1
    tableRow:      140, tableCol: 1,
  },
  {
    title:         "Bell Canada - CAD Re-Offer Yield Curve",
    cols:          [4, 6, 8, 10, 12],
    yLabel:        "Yield (%)",
    isPct:         true,
    chartAnchorRow: 1, chartAnchorCol: 17,  // Q1
    tableRow:      140, tableCol: 10,
  },
  {
    title:         "Bell Canada - USD New Issue Spread Curve (bps)",
    cols:          [13, 15, 17, 19, 21],
    yLabel:        "Spread (bps)",
    isPct:         false,
    chartAnchorRow: 33, chartAnchorCol: 1,  // A33
    tableRow:      170, tableCol: 1,
  },
  {
    title:         "Bell Canada - USD Re-Offer Yield Curve",
    cols:          [14, 16, 18, 20, 22],
    yLabel:        "Yield (%)",
    isPct:         true,
    chartAnchorRow: 33, chartAnchorCol: 17, // Q33
    tableRow:      170, tableCol: 10,
  },
];

const TENORS = ["3Y", "5Y", "7Y", "10Y", "30Y"];

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

function main(workbook: ExcelScript.Workbook, dataJson: string): string {
  let data: { [key: string]: string | number | null };
  try {
    data = JSON.parse(dataJson);
  } catch (e) {
    return `ERROR: dataJson is not valid JSON. ${e}`;
  }

  try {
    appendRow(workbook, data);
  } catch (e) {
    return `ERROR appending row: ${e}`;
  }

  try {
    updateCharts(workbook);
  } catch (e) {
    return `ERROR updating charts: ${e}`;
  }

  return `OK: wrote ${data["bank"]} data for ${data["date"]}.`;
}

// ---------------------------------------------------------------------------
// appendRow – mirrors excel_writer.append_row()
// ---------------------------------------------------------------------------

function appendRow(
  workbook: ExcelScript.Workbook,
  data: { [key: string]: string | number | null }
): void {
  const ws = workbook.getWorksheet("Pricing");
  if (!ws) throw new Error('Sheet "Pricing" not found.');

  // Find the first empty row starting from row 2 (row index 1 = row 2 in UI).
  const usedRange = ws.getUsedRange();
  const lastRow = usedRange ? usedRange.getRowCount() : 1; // 1-indexed count
  let nextRow = lastRow + 1; // default: one past last used row

  // Walk down from row 2 to find the actual first empty cell in column A.
  for (let r = 2; r <= lastRow + 1; r++) {
    const cell = ws.getCell(r - 1, 0); // 0-indexed
    if (cell.getValue() === null || cell.getValue() === "") {
      nextRow = r;
      break;
    }
  }

  const rowIdx = nextRow - 1; // convert to 0-based for ExcelScript API

  // Column A (index 0): date – stored as a date value, formatted YYYY-MM-DD.
  const dateCell = ws.getCell(rowIdx, 0);
  dateCell.setValue(data["date"] as string);
  dateCell.setNumberFormat("YYYY-MM-DD");
  dateCell.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  // Column B (index 1): bank name.
  const bankCell = ws.getCell(rowIdx, 1);
  bankCell.setValue(data["bank"] as string);
  bankCell.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  // Metric columns (3–30, 1-indexed → 2–29, 0-indexed).
  for (const [colStr, key] of Object.entries(COLUMN_MAP)) {
    const col1 = Number(colStr);       // 1-indexed
    const colIdx = col1 - 1;           // 0-indexed
    const value = data[key];
    if (value === null || value === undefined) continue;

    const cell = ws.getCell(rowIdx, colIdx);
    cell.setValue(value as number);
    cell.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    if (PCT_COLUMNS.has(col1)) {
      cell.setNumberFormat("0.000%");
    }
  }
}

// ---------------------------------------------------------------------------
// updateCharts – mirrors excel_writer.update_charts()
// ---------------------------------------------------------------------------

function updateCharts(workbook: ExcelScript.Workbook): void {
  const wsPricing = workbook.getWorksheet("Pricing");
  const wsCharts  = workbook.getWorksheet("Summary Charts");
  if (!wsPricing) throw new Error('Sheet "Pricing" not found.');
  if (!wsCharts)  throw new Error('Sheet "Summary Charts" not found.');

  // --- Clear all existing charts from Summary Charts ----------------------
  for (const chart of wsCharts.getCharts()) {
    chart.delete();
  }

  // --- Collect data rows from Pricing -------------------------------------
  // Only contiguous filled rows starting from row 2.
  const usedRange = wsPricing.getUsedRange();
  const totalRows = usedRange ? usedRange.getRowCount() : 1;

  interface RowEntry { row: number; date: string; bank: string }
  const allRows: RowEntry[] = [];

  for (let r = 2; r <= totalRows; r++) {
    const dateVal = wsPricing.getCell(r - 1, 0).getValue();
    const bankVal = wsPricing.getCell(r - 1, 1).getValue();
    if (dateVal && bankVal) {
      allRows.push({ row: r, date: String(dateVal), bank: String(bankVal) });
    } else {
      break; // stop at first gap, same as python version
    }
  }

  if (allRows.length === 0) return;

  // --- Keep only the most recent row per bank ------------------------------
  const latestPerBank: { [bank: string]: RowEntry } = {};
  for (const entry of allRows) {
    if (
      !(entry.bank in latestPerBank) ||
      entry.date > latestPerBank[entry.bank].date
    ) {
      latestPerBank[entry.bank] = entry;
    }
  }
  const rows = Object.values(latestPerBank).sort((a, b) =>
    a.date < b.date ? -1 : a.date > b.date ? 1 : 0
  );

  // --- Build one staging table + one chart per config ---------------------
  for (const cfg of CHART_CONFIGS) {
    const sr = cfg.tableRow;     // 1-indexed staging table header row
    const sc = cfg.tableCol;     // 1-indexed staging table header col

    // Write header row: "Bank", then tenor labels.
    setCell(wsCharts, sr, sc, "Bank");
    for (let j = 0; j < TENORS.length; j++) {
      setCell(wsCharts, sr, sc + 1 + j, TENORS[j]);
    }

    // Write one data row per bank/date.
    const yValues: number[] = [];
    for (let i = 0; i < rows.length; i++) {
      const entry = rows[i];
      const label = `${entry.bank} (${entry.date})`;
      setCell(wsCharts, sr + 1 + i, sc, label);

      for (let j = 0; j < cfg.cols.length; j++) {
        const pricingCol = cfg.cols[j]; // 1-indexed Pricing sheet column
        const raw = wsPricing.getCell(entry.row - 1, pricingCol - 1).getValue();
        const val = typeof raw === "number" ? raw : null;
        const cell = wsCharts.getCell(sr + i, sc + j); // 0-indexed
        if (val !== null) {
          cell.setValue(val);
          yValues.push(val);
          if (cfg.isPct) cell.setNumberFormat("0.00%");
        }
      }
    }

    // --- Build chart -------------------------------------------------------
    // Data range: header row + one row per bank (labels col excluded from chart).
    const dataRowCount = rows.length;

    // Category range: the tenor header row (cols sc+1 .. sc+5), 1-indexed.
    const catRange = wsCharts.getRangeByIndexes(
      sr - 1,          // 0-based row of header
      sc,              // 0-based col of first tenor label
      1,               // 1 row
      TENORS.length    // 5 cols
    );

    // Chart data range: the full block including labels (col sc) + values.
    // We add series manually below for cleaner label control.
    const chart = wsCharts.addChart(
      ExcelScript.ChartType.line,
      wsCharts.getRangeByIndexes(sr - 1, sc - 1, 1 + dataRowCount, 1 + TENORS.length)
    );

    chart.setName(cfg.title);
    chart.getTitle().setText(cfg.title);

    // Position and size (roughly equivalent to 24cm × 14cm in openpyxl).
    chart.setTop((cfg.chartAnchorRow - 1) * 20);
    chart.setLeft((cfg.chartAnchorCol - 1) * 64);
    chart.setWidth(720);   // ~24cm at 96 dpi
    chart.setHeight(420);  // ~14cm at 96 dpi

    // Axes.
    const xAxis = chart.getAxes().getCategoryAxis();
    xAxis.getTitle().setText("Tenor");
    xAxis.getTitle().setVisible(true);

    const yAxis = chart.getAxes().getValueAxis();
    yAxis.getTitle().setText(cfg.yLabel);
    yAxis.getTitle().setVisible(true);
    yAxis.setNumberFormat(cfg.isPct ? "0.00%" : "0");

    // Y-axis scaling: match _major_unit() logic from excel_writer.py.
    if (yValues.length > 0) {
      const yMin = Math.min(...yValues);
      const yMax = Math.max(...yValues);
      const span = Math.max(yMax - yMin, 1e-9);
      const major = majorUnit(span, cfg.isPct);
      yAxis.setMajorUnit(major);
      yAxis.setMinimumValue(Math.max(0, yMin - major));
      yAxis.setMaximumValue(yMax + major);
    }

    // Series: one per bank. Replace the auto-generated series with correctly
    // labelled ones that reference the staging table rows.
    // Remove all auto-added series first.
    const existingSeries = chart.getSeries();
    for (const s of existingSeries) {
      s.delete();
    }

    for (let i = 0; i < rows.length; i++) {
      const entry = rows[i];
      const series = chart.addChartSeries();
      series.setName(`${entry.bank} (${entry.date})`);

      // Values range: staging table row i+1 (0-based: sr + i), cols sc+1..sc+5.
      const valRange = wsCharts.getRangeByIndexes(
        sr + i,        // 0-based: row after header + offset
        sc,            // 0-based: first value col
        1,
        TENORS.length
      );
      series.setValues(valRange);
      series.setXAxisValues(catRange);
      series.setMarkerStyle(ExcelScript.ChartMarkerStyle.circle);
      series.setMarkerSize(7);
      series.setSmooth(false);
    }

    // Legend on the right.
    chart.getLegend().setVisible(true);
    chart.getLegend().setPosition(ExcelScript.ChartLegendPosition.right);
    chart.getLegend().setOverlay(false);
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Write a value to a 1-indexed (row, col) cell and centre-align it. */
function setCell(
  ws: ExcelScript.Worksheet,
  row: number,
  col: number,
  value: string | number
): void {
  const cell = ws.getCell(row - 1, col - 1);
  cell.setValue(value);
  cell.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
}

/**
 * Choose a readable y-axis major unit.
 * Mirrors excel_writer._major_unit() exactly.
 */
function majorUnit(span: number, isPct: boolean): number {
  const candidates = isPct
    ? [0.00025, 0.0005, 0.001, 0.002, 0.0025, 0.005, 0.01]
    : [0.5, 1.0, 2.0, 2.5, 5.0, 10.0, 20.0, 25.0, 50.0];

  const target = Math.max(span / 12.0, candidates[0]);
  for (const step of candidates) {
    if (step >= target) return step;
  }
  return candidates[candidates.length - 1];
}
