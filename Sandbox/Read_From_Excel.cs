using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab

namespace Sandbox
{
    public class Read_From_Excel
    {
        public static List<string> getExcelFile(string path, StreamWriter writer)
        {

            List<string> chemia = new List<string>();


            //Create COM Objects. Create a COM object for everything that is referencedddd
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;


            Excel.Range currentFind = null;
            Excel.Range Chempol = xlWorksheet.UsedRange;
            currentFind = Chempol.Find("Lp.", Type.Missing,
                Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
                Type.Missing, Type.Missing);

            if (currentFind == null) {
                using (StreamWriter sw = writer) {
                    writer.WriteLine(path + "  nie znaleziono 'Lp.'");
                        }
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                return new List<string>();

            }

            int m = currentFind.Row;
            int LpStart = m;
            int n = currentFind.Column;

            string rowPosition = (string)(xlWorksheet.Cells[m, n] as Excel.Range).Value2;

            while (rowPosition != "1")
            {
                m += 1;
                rowPosition = ToStr((xlWorksheet.Cells[m, n] as Excel.Range).Value2);

            }

            string columnPosition = ToStr((xlWorksheet.Cells[m, n] as Excel.Range).Value2);

            while (columnPosition != "Rodzaj farby/rozpuszczalnika/ aktywatora/lakieru")
            {
                n += 1;
                columnPosition = ToStr((xlWorksheet.Cells[LpStart, n] as Excel.Range).Value2);

                if (n > 100) {

                    using (StreamWriter sw = writer)
                    {
                        writer.WriteLine(path + "  nie znaleziono 'Rodzaj farby/rozpuszczalnika/ aktywatora/lakieru'");
                    }
                    Marshal.ReleaseComObject(xlRange);
                    Marshal.ReleaseComObject(xlWorksheet);
                    xlWorkbook.Close();
                    Marshal.ReleaseComObject(xlWorkbook);
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);
                    return new List<string>();
                }

            }

            //int m = currentFind.Row + 2;
            //int n = currentFind.Column + 2;


            for (int k = 0; k < 10; k++)
            {
                var cellValue = ToStr((xlWorksheet.Cells[m + k, n] as Excel.Range).Value2);
                chemia.Add(cellValue);

            }




            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            //xlApp = null;
            //xlRange = null;
            //xlWorkbook = null;
            //xlWorksheet = null;
            return chemia;
        }

        public static string ToStr(object readField)
        {
            if ((readField != null))
            {
                if (readField.GetType() != typeof(System.DBNull))
                {
                    return Convert.ToString(readField);
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }
    }


}