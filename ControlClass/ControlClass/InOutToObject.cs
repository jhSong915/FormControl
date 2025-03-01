using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace ControlClass
{
    public class InOutToObject
    {
        public enum ExcelType
        {
            EXCEL_EXPORT,
            EXCEL_IMPORT,
            CSV_EXPORT
        }

        Excel.Application excelApp = null;
        Excel.Workbook wb = null;
        Excel.Worksheet ws = null;
        Excel.Worksheet xlSheet = null;
        Excel.Range xlRange = null;

        public string ExcelExport(string SaveFilePath, DataTable dt, ExcelType excelType)
        {
            string resultMsg = string.Empty;

            switch(excelType)
            {
                case ExcelType.EXCEL_EXPORT:
                    resultMsg = ExportExcel_API(SaveFilePath, dt);
                    break;
                case ExcelType.CSV_EXPORT:
                    resultMsg = ExportExcel_CSV(SaveFilePath, dt);
                    break;
            }
            return resultMsg;
        }

        public DataTable ExcelImport(string SaveFilePath, ExcelType excelType, string ReporType, int HeaderLine, int ColumnStart)
        {
            DataTable resultDt = new DataTable();

            switch (excelType)
            {
                case ExcelType.EXCEL_IMPORT:
                    resultDt = ImportExcel_API(SaveFilePath, ReporType, HeaderLine, ColumnStart);
                    break;
            }
            return resultDt;
        }

        private DataTable ImportExcel_API(string LoadFilePath, string ReporType, int HeaderLine, int ColumnStart)
        {
            DataTable resultDt = new DataTable();

            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;
                wb = excelApp.Workbooks.Open(LoadFilePath);
                xlSheet = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.Item[0];
                xlRange = xlSheet.UsedRange;

                int cl = xlRange.Columns.Count;
                int rowcount = xlRange.Rows.Count; ;
                //create the header of table
                for (int j = ColumnStart; j <= cl; j++)
                {
                    resultDt.Columns.Add(Convert.ToString(xlRange.Cells[HeaderLine, j].Value2), typeof(string));
                }
                //filling the table from  excel file                
                for (int i = HeaderLine + 1; i <= rowcount; i++)
                {
                    DataRow dr = resultDt.NewRow();
                    for (int j = ColumnStart; j <= cl; j++)
                    {
                        dr[j - ColumnStart] = Convert.ToString(xlRange.Cells[i, j].Value2);
                    }
                    resultDt.Rows.InsertAt(dr, resultDt.Rows.Count + 1);
                }
            }
            catch (Exception ex)
            {
                resultDt.Reset();
                resultDt.Rows[0]["Status"] = "Error";
                resultDt.Rows[0]["Message"] = $"Excel Import Failed : {ex}";
            }
            finally
            {
                ReleaseExcelObject(ws); ReleaseExcelObject(wb); ReleaseExcelObject(excelApp);
            }

            return resultDt;
        }

        private string ExportExcel_API(string SaveFilePath, DataTable dt)
        {
            string resultMsg = string.Empty;
            Excel.Application excelApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                excelApp = new Excel.Application();
                wb = excelApp.Workbooks.Add();
                // Excel First Worksheet Read
                ws = wb.Worksheets.get_Item(1) as Excel.Worksheet;

                int iRow = 1; foreach (DataRow dr in dt.Rows)
                {
                    int iColumn = 1; foreach (var vitem in dr.ItemArray)
                    {
                        ws.Cells[iRow, iColumn] = vitem;
                        iColumn++;                    
                    }
                    iRow++;
                }
                // Save Excel File
                wb.SaveAs(SaveFilePath, Excel.XlFileFormat.xlWorkbookNormal);
                wb.Close(true);
                excelApp.Quit();

                resultMsg = "Excel Export Success!";
            }
            catch (Exception ex)
            {
                resultMsg = $"Excel Export Failed : {ex}";
            }
            finally
            {
                ReleaseExcelObject(ws); ReleaseExcelObject(wb); ReleaseExcelObject(excelApp);
            }
            return resultMsg;
        }

        private string ExportExcel_CSV(string SaveFilePath, DataTable dt)
        {
            string resultMsg = string.Empty;
            try
            {
                var lines = new List<string>();
                string[] columnNames = dt.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray();
                var header = string.Join(",", columnNames);
                lines.Add(header);
                var valueLines = dt.AsEnumerable().Select(row => string.Join(",", row.ItemArray));
                lines.AddRange(valueLines);
                File.WriteAllLines(SaveFilePath, lines, Encoding.UTF8);
                return "Excel Export Success!";            
            }
            catch (Exception ex)
            {
                return $"Excel Export Failed : {ex}";
            }
        }

        private void ReleaseExcelObject(object obj) 
        { 
            try
            { 
                if (obj != null) 
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj); obj = null;
                } 
            } 
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            } 
            finally 
            { 
                GC.Collect(); 
            }
        }
    }
}
