using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace InventorModelStateCleanup
{
    // ================================================================
    // PART 2 OF 3: EXCEL READER
    // Reads Model State names from the Excel file.
    // ================================================================
    static class ExcelReader
    {
        // ============================================================
        // BLOCK 6: EXCEL SETTINGS
        // ============================================================

        const int StartRow = 10; // first data row
        const int Column   = 2;  // column B

        // ============================================================
        // BLOCK 7: READ NAMES
        // Reads column B from row 10 until the first empty cell.
        // ============================================================

        public static List<string> ReadModelStateNames(string excelPath)
        {
            List<string> names = new List<string>();

            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                workbook = excelApp.Workbooks.Open(excelPath);
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];

                for (int row = StartRow; ; row++)
                {
                    Excel.Range cell = null;

                    try
                    {
                        cell = (Excel.Range)worksheet.Cells[row, Column];
                        string value = Convert.ToString(cell.get_Value(Type.Missing));

                        if (string.IsNullOrWhiteSpace(value))
                            break;

                        value = value.Trim();
                        names.Add(value);
                        Console.WriteLine("Read Model State: " + value);
                    }
                    finally
                    {
                        ReleaseComObject(cell);
                    }
                }
            }
            finally
            {
                if (workbook != null)
                {
                    try { workbook.Close(false); } catch { }
                }
                if (excelApp != null)
                {
                    try { excelApp.Quit(); } catch { }
                }

                ReleaseComObject(worksheet);
                ReleaseComObject(workbook);
                ReleaseComObject(excelApp);
            }

            return names;
        }

        // ============================================================
        // BLOCK 8: RELEASE COM OBJECTS
        // ============================================================

        static void ReleaseComObject(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                try { Marshal.ReleaseComObject(obj); }
                catch { }
            }
        }
    }
}
