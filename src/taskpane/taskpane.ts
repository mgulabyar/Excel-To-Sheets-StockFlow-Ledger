declare const Excel: any;
export async function insertText(text: string) {
  try {
    await Excel.run(async (context: { workbook: { worksheets: { getActiveWorksheet: () => any; }; }; sync: () => any; }) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A1");
      range.values = [[text]];
      range.format.autofitColumns();
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}
