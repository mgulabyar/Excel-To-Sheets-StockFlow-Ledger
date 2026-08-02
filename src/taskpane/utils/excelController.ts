// ============================================================================
// excelController.ts - Dedicated Office.js Helper Layer
// Fixes: all Office.js logic was previously inline inside App.tsx
// ============================================================================

declare const Excel: any;

const RISK_HIGHLIGHT_COLOR = "#FEE2E2";
const CONFLICT_HIGHLIGHT_COLOR = "#FEF3C7"; // amber - distinct from stockout red

/**
 * Registers a selection-change listener on the active worksheet.
 * Returns a disposer so the caller (React useEffect) can clean it up.
 */
export async function bindSelectionHandler(
  onSelectionChanged: (itemCode: string, quantity: string, rowIndex: number) => void
): Promise<() => void> {
  let handlerRef: any = null;

  await Excel.run(async (context: any) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    const callback = async () => {
      await Excel.run(async (innerContext: any) => {
        const innerSheet = innerContext.workbook.worksheets.getActiveWorksheet();
        const selectedRange = innerContext.workbook.getSelectedRange();
        selectedRange.load("rowIndex");
        await innerContext.sync();

        const codeRange = innerSheet.getRangeByIndexes(selectedRange.rowIndex, 0, 1, 1);
        const qtyRange = innerSheet.getRangeByIndexes(selectedRange.rowIndex, 2, 1, 1);
        codeRange.load("values");
        qtyRange.load("values");
        await innerContext.sync();

        if (codeRange.values[0][0]) {
          onSelectionChanged(
            String(codeRange.values[0][0]).trim(),
            String(qtyRange.values[0][0]).trim(),
            selectedRange.rowIndex
          );
        }
      });
    };

    handlerRef = sheet.onSelectionChanged.add(callback);
    await context.sync();
  });

  return () => {
    // Office.js event removal - handlerRef is a subscription object
    if (handlerRef && handlerRef.remove) {
      handlerRef.remove();
    }
  };
}

/**
 * Reads the full used range as a 2D array of values.
 */
export async function readUsedRange(): Promise<any[][]> {
  let values: any[][] = [];
  await Excel.run(async (context: any) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    usedRange.load("values");
    await context.sync();
    values = usedRange.values;
  });
  return values;
}

/**
 * Applies stockout-risk highlight or clears it, for a given item row.
 */
export async function applyStockoutHighlight(itemCode: string, isAtRisk: boolean): Promise<void> {
  await Excel.run(async (context: any) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    usedRange.load("values");
    await context.sync();

    const dataValues = usedRange.values;
    for (let i = 1; i < dataValues.length; i++) {
      if (String(dataValues[i][0]).trim() === itemCode.trim()) {
        const targetRowRange = sheet.getRow(i);
        const visualDataGrid = targetRowRange.getResizedRange(0, 4);
        if (isAtRisk) {
          visualDataGrid.format.fill.color = RISK_HIGHLIGHT_COLOR;
        } else {
          visualDataGrid.format.fill.clear();
        }
        break;
      }
    }
    await context.sync();
  });
}


export async function applyConflictHighlight(itemCode: string): Promise<void> {
  await Excel.run(async (context: any) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    usedRange.load("values");
    await context.sync();

    const dataValues = usedRange.values;
    for (let i = 1; i < dataValues.length; i++) {
      if (String(dataValues[i][0]).trim() === itemCode.trim()) {
        const targetRowRange = sheet.getRow(i);
        const visualDataGrid = targetRowRange.getResizedRange(0, 4);
        visualDataGrid.format.fill.color = CONFLICT_HIGHLIGHT_COLOR;
        break;
      }
    }
    await context.sync();
  });
}


export async function writeRowUpdate(
  itemCode: string,
  newQuantity: number,
  statusLabel: string
): Promise<{ previousQuantity: number | null; rowFound: boolean }> {
  let previousQuantity: number | null = null;
  let rowFound = false;

  await Excel.run(async (context: any) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    usedRange.load("values");
    await context.sync();

    const dataValues = usedRange.values;
    for (let i = 1; i < dataValues.length; i++) {
      if (String(dataValues[i][0]).trim() === itemCode.trim()) {
        rowFound = true;
        previousQuantity = Number(dataValues[i][2]) || null;

        sheet.getRangeByIndexes(i, 2, 1, 1).values = [[newQuantity]];
        sheet.getRangeByIndexes(i, 4, 1, 1).values = [[statusLabel]];

        const targetRowRange = sheet.getRow(i);
        const visualDataGrid = targetRowRange.getResizedRange(0, 4);
        visualDataGrid.format.fill.clear();
        break;
      }
    }
    await context.sync();
  });

  return { previousQuantity, rowFound };
}


export async function appendAuditRowToWorkbook(row: (string | number)[]): Promise<void> {
  await Excel.run(async (context: any) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();

    let auditSheet = worksheets.items.find((s: any) => s.name === "AuditLog");
    if (!auditSheet) {
      auditSheet = worksheets.add("AuditLog");
      auditSheet.getRange("A1:F1").values = [
        ["ID", "Item Code", "Action", "Before", "After", "Timestamp"],
      ];
      auditSheet.visibility = "Hidden";
    }

    const usedRange = auditSheet.getUsedRange();
    usedRange.load("rowCount");
    await context.sync();

    const nextRow = usedRange.rowCount;
    auditSheet.getRangeByIndexes(nextRow, 0, 1, row.length).values = [row];
    await context.sync();
  });
}