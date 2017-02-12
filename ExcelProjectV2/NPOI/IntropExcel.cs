using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace ExcelProjectV2.NPOI
{
    public class IntropExcel
    {
        Excel.Application app = new Excel.Application();
        Excel.Workbook curWorkBook = null;
        Excel.Workbook destWorkbook = null;
        Excel.Worksheet workSheet = null;
        Excel.Worksheet newWorksheet = null;
        Object defaultArg = Type.Missing;



        public string CreateTempFile(string path, List<int> sheets)
        {
            curWorkBook = app.Workbooks.Open(path);
            // app.Visible = true;
            // 
            app.DisplayAlerts = false;



            var filPath = Path.GetFullPath(path);
            var ext = Path.GetExtension(filPath).ToLower();
            var fileName = Path.GetFileNameWithoutExtension(filPath);
            var fileNamewithExt = Path.GetFileName(filPath);
            var filePathwithoutFileName = filPath.Substring(0, filPath.Length - fileNamewithExt.Length);
            var tempFilePath = filePathwithoutFileName + "Temp" + fileNamewithExt;
            var fileToCopy = Path.GetFullPath("Temp" + ext);
            //File.Copy( fileToCopy ,tempFilePath,true);
            File.Copy(fileToCopy, tempFilePath, true);
            destWorkbook = app.Workbooks.Add(tempFilePath);
            try
            {
                foreach (int sheet1 in sheets)
                {
                    Excel.Worksheet sheet = curWorkBook.Sheets[sheet1+1];
                    sheet.UsedRange.Copy(defaultArg);
                    var osheet = destWorkbook.Sheets.Add();
                    osheet.Name = curWorkBook.Sheets[sheet1+1].Name;
                    //  newWorksheet = (Excel.Worksheet)destWorkbook.Worksheets.Add(defaultArg, defaultArg, defaultArg, defaultArg);
                    osheet.UsedRange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                }

            }
            catch (Exception exc)
            {
                System.Windows.Forms.MessageBox.Show(exc.Message);
            }
            finally
            {


                if (destWorkbook != null)
                {
                    if (ext == ".xlsx")
                    {
                       
                        destWorkbook.SaveAs(tempFilePath, Excel.XlFileFormat.xlOpenXMLWorkbook,
        System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false,
        Excel.XlSaveAsAccessMode.xlNoChange, false, false, System.Reflection.Missing.Value,
        System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                        destWorkbook.Close(defaultArg, defaultArg, defaultArg);
                    }
                    else if (ext == ".xls")
                    {
                        destWorkbook.SaveAs(tempFilePath, Excel.XlFileFormat.xlWorkbookDefault,
        System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false,
        Excel.XlSaveAsAccessMode.xlNoChange, false, false, System.Reflection.Missing.Value,
        System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                        destWorkbook.Close(defaultArg, defaultArg, defaultArg);
                    }
                }
                app.Quit();
                var proc = Process.GetProcessesByName("EXCEL");
                foreach (var process in proc)
                {
                    process.Kill();
                }
            }
            return tempFilePath;
        }
    }
}
