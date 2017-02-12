using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProjectV2.NPOI
{
    using System.Data;
    using System.IO;
    using System.Runtime.Remoting.Contexts;
    using System.Windows.Forms;
    using System.Xml.Xsl;

    using ExcelProjectV2.Model;

    using global::NPOI.HSSF.UserModel;
    using global::NPOI.OpenXmlFormats.Shared;
    using global::NPOI.SS.UserModel;
    using global::NPOI.SS.Util;
    using global::NPOI.XSSF.UserModel;

    class NPOIFactory
    {
        #region  DeclearVaribles




        private IWorkbook hssWB = null;

        private IFormulaEvaluator helpr = null;

        private DataTable inputTable = null;

        private DataTable AtmTable = null;

        private DataTable PrintAtmTable = null;

        private DataTable CashTable = null;

        private DataFormatter df = null;

        private NpoiStyle style = null;

        #endregion

        #region Constructor

        public NPOIFactory(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("من فضلك أختار ملف ");
            }
            string ext = Path.GetExtension(filePath).ToLower();
            if (ext != ".xls" && ext != ".xlsx")
            {
                MessageBox.Show("عفوا لا يمكن التعامل مع صيغة هذا الملف ");
            }

            try
            {
                using (Stream file = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
                {
                    if (ext == ".xls")
                    {
                        // hssWB = new HSSFWorkbook(file);
                        hssWB = WorkbookFactory.Create(file);

                        //this.helpr = new HSSFFormulaEvaluator(this.hssWB);
                        this.helpr = hssWB.GetCreationHelper().CreateFormulaEvaluator();
                        file.Close();
                    }
                    else if (ext == ".xlsx")
                    {
                        // this.hssWB = new XSSFWorkbook(file);
                        hssWB = WorkbookFactory.Create(file);
                        //   this.helpr = new XSSFFormulaEvaluator(this.hssWB);
                        this.helpr = hssWB.GetCreationHelper().CreateFormulaEvaluator();
                        file.Close();

                    }
                    this.FilePath = filePath;
                    this.df = new DataFormatter();
                    this.inputTable = new DataTable();
                    this.AtmTable = new DataTable();
                    this.PrintAtmTable = new DataTable();
                    this.CashTable = new DataTable();
                    this.style= new NpoiStyle();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        string FilePath;

        #endregion


        #region read excel

        public List<string> GetSheetsName()
        {
            List<string> SheetsName = new List<string>();

            try
            {
                for (int i = 0; i < hssWB.NumberOfSheets; i++)
                {
                    SheetsName.Add(hssWB.GetSheetAt(i).SheetName);
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

            return SheetsName;

        }


        public DataTable ReadExcel(List<string> sheets, int startRow, string paymentType)
        {
            try
            {
                foreach (string sheet in sheets)
                {
                    ISheet currentSheet = this.hssWB.GetSheet(sheet);
                    IRow codeRow = currentSheet.GetRow(startRow);
                    IRow nameRow = currentSheet.GetRow(startRow + 1);
                    var header = GenertareHeader(codeRow, nameRow);
                    ReadOriginalSheet(currentSheet, header, startRow + 2);
                  
                 
                }
                SplitTable(this.inputTable, paymentType);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

            return this.inputTable;
        }

        private void SplitTable(DataTable dataTable, string paymentType)
        {
            this.AtmTable.Clear();
            this.CashTable.Clear();
            this.PrintAtmTable.Clear();


            var codeIndex = dataTable.Columns["3003-الكود"].Ordinal;
            int NameIndex = dataTable.Columns["3001-الأسم"].Ordinal;
            int NetIndex = dataTable.Columns["3333-الصافى"].Ordinal;
            ERPEntities context = new ERPEntities();

            foreach (DataRow row in dataTable.Rows)
            {
                var code = row.ItemArray[codeIndex].ToString();
                if (code == "0" || string.IsNullOrEmpty(code))
                {
                    DataRow cashDr = this.CashTable.NewRow();
                    cashDr[0] = row.ItemArray[NameIndex];
                    cashDr[1] = row.ItemArray[NetIndex];
                    cashDr[2] = string.Empty;
                    this.CashTable.Rows.Add(cashDr);
                }
                else
                {
                    DataRow atmdr = this.AtmTable.NewRow();
                    DataRow printRow = this.PrintAtmTable.NewRow();
                    int i = 0;
                    int code2 = 0;
                    bool convert = int.TryParse(row.ItemArray[codeIndex].ToString(), out code2);
                    Employee emp = null;
                    if (convert)
                    {
                        emp = context.Employees.Find(code2);
                        printRow[0] = emp.NationalId;
                        printRow[1] = paymentType;
                        printRow[2] = string.Empty;
                        printRow[3] = emp.PositionName;
                        printRow[4] = emp.Code;
                        printRow[5] = emp.Name;
                        printRow[6] = row.ItemArray[NetIndex];

                        foreach (var cell in row.ItemArray)
                        {


                            if (cell != null)
                            {

                                if (i == NameIndex)
                                {
                                    atmdr[i] = emp.Name;
                                }
                                else
                                {
                                    atmdr[i] = cell.ToString();
                                }
                            }
                            else
                            {
                                atmdr[i] = string.Empty;
                            }
                            i++;
                        }
                        this.AtmTable.Rows.Add(atmdr);
                        this.PrintAtmTable.Rows.Add(printRow);
                    }
                }
            }

        }

        private DataTable ReadOriginalSheet(ISheet sheet, HeaderList header, int startRow)
        {
            try
            {
                if (this.inputTable.Columns.Count == 0)
                {

                    foreach (HeaderContent headerdata in header.headerContent)
                    {
                        this.inputTable.Columns.Add(headerdata.HeaderCode + "-" + headerdata.HeaderName);
                        this.AtmTable.Columns.Add(headerdata.HeaderCode + "-" + headerdata.HeaderName);
                    }

                    this.CashTable.Columns.Add("الأسم");
                    this.CashTable.Columns.Add("المبلغ");
                    this.CashTable.Columns.Add("التوقيع");

                    this.PrintAtmTable.Columns.Add(new DataColumn("الرقم القومى"));
                    this.PrintAtmTable.Columns.Add(new DataColumn("نوع المدفوعه"));
                    this.PrintAtmTable.Columns.Add(new DataColumn("القطاع"));
                    this.PrintAtmTable.Columns.Add(new DataColumn("الإدارة"));
                    this.PrintAtmTable.Columns.Add(new DataColumn("رقم الموظف بجهته الأصلية"));
                    this.PrintAtmTable.Columns.Add(new DataColumn("الاسم"));
                    this.PrintAtmTable.Columns.Add(new DataColumn("المرتب"));
                }

                //TODO Message Code Not Present
                int? nameIndex = null;
                int? netIndex = null;
                if (header.headerContent.FirstOrDefault(x => x.HeaderCode == 3001) != null)
                {
                   nameIndex = (int?)header.headerContent.FirstOrDefault(x => x.HeaderCode == 3001).ColIndex;
                }
               else                 {
              
                    throw new Exception("3001 تأكد من وجود كود الأسم");
                }
                if (header.headerContent.FirstOrDefault(x => x.HeaderCode == 3333) != null)
                { 
                   netIndex = header.headerContent.FirstOrDefault(x => x.HeaderCode == 3333).ColIndex;}
             else
                {
                    throw new Exception("3333 تأكد من وجود كود الصافى");
                }
                int lastRow = sheet.LastRowNum+1;

                for (int i = startRow; i < lastRow; i++)
                {
                    IRow currentRow = sheet.GetRow(i);
                  //  if (currentRow != null&&!string.IsNullOrEmpty( currentRow.Cells[nameIndex].StringCellValue)&& currentRow.Cells[nameIndex].StringCellValue!="0")
                    if (currentRow != null)
                    {

                        var namecell = currentRow.GetCell(nameIndex.Value);
                        var netecell = currentRow.GetCell(netIndex.Value);
                        string nameval = string.Empty;
                        decimal income = 0;
                        decimal outcome = 0;
                        if (namecell!=null )
                        {
                            if (namecell.CellType == CellType.Formula)
                            {
                                this.helpr.EvaluateInCell(namecell);
                            }
                       
                         nameval = this.df.FormatCellValue(namecell);
                        }
                        string netVal = string.Empty;
                        int datarowindex = 0;
                        if (nameval != "الأجمالى"&& nameval != "الأسم" && !string.IsNullOrEmpty(nameval) && nameval != "0")
                        {

                            DataRow dr = this.inputTable.NewRow();
                            foreach (var cell in header.headerContent)
                            {

                                ICell currentCell = null;
                                if (currentRow.GetCell(cell.ColIndex) == null)
                                {
                                    currentCell = currentRow.CreateCell(cell.ColIndex);
                                    currentCell.SetCellValue(  string.Empty);
                                }

                                currentCell = currentRow.GetCell(cell.ColIndex);
                                if (currentCell.CellType == CellType.Formula)
                                {
                                    this.helpr.EvaluateInCell(netecell);

                                    this.helpr.EvaluateInCell(currentCell);
                                }
                                string cellVal = this.df.FormatCellValue(currentCell);
                                if (cellVal == string.Empty)
                                {
                                    cellVal = "0";
                                }
                                if (cell.HeaderCode.ToString().StartsWith("1"))
                                {
                                    income += decimal.Parse(cellVal);
                                }
                                if (cell.HeaderCode.ToString().StartsWith("2"))
                                {
                                    outcome += decimal.Parse(cellVal);
                                }
                                if (cell.HeaderCode.ToString() == "3333")
                                {
                                     netVal = this.df.FormatCellValue(netecell);
                                   
                                }
                                dr[datarowindex] = cellVal;
                                datarowindex++;
                            }
                            if (decimal.Parse(netVal) != (income - outcome))
                            {
                                //MessageBox.Show(
                                //    "يوجد خطأ فى أجمالى " + currentRow.GetCell(nameIndex.Value).StringCellValue
                                //    + "الصافى يجب ان يكون " + (income - outcome).ToString() + "بدلا من "
                                //    + netVal.ToString());


                            throw new Exception("يوجد خطأ فى أجمالى " + currentRow.GetCell(nameIndex.Value).StringCellValue
                                  + "الصافى يجب ان يكون " + (income - outcome).ToString() + "بدلا من "
                                    + netVal.ToString());
                            }
                            this.inputTable.Rows.Add(dr);
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

            return this.inputTable;


        }

        public HeaderList GenertareHeader(IRow codeRow, IRow nameRow)
        {
            HeaderList header = new HeaderList();
            ERPEntities context = new ERPEntities();
            try
            {
                byte counter = 0;
                foreach (ICell cell in codeRow.Cells)
                {
                    var cellVal = this.df.FormatCellValue(cell);
                    var headerName = this.df.FormatCellValue(nameRow.GetCell(cell.ColumnIndex));

                    if (!string.IsNullOrEmpty(cellVal))
                    {
                        int code = 0;
                        bool convert = int.TryParse(cellVal, out code);
                        if (convert)
                        {
                            var checkcode = context.Accounts.FirstOrDefault(x => x.Id == code);
                            if (checkcode == null)
                            {
                                //  MessageBox.Show("تأكد من الكود رقم " + code);
                                throw new Exception("تأكد من الكود رقم " + code);
                            }
                            string type = string.Empty;
                            if (code.ToString().StartsWith("1"))
                            {
                                type = "Credit";
                            }
                            if (code.ToString().StartsWith("2"))
                            {
                                type = "Debit";
                            }
                            if (code.ToString().StartsWith("3"))
                            {
                                type = "Def";
                            }

                            header.headerContent.Add(
                                new HeaderContent()
                                {
                                    ColIndex = cell.ColumnIndex,
                                    HeaderCode = code,
                                    HeaderName = headerName,
                                    HeaderType = type
                                });
                            counter++;
                        }

                    }

                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            return header;
        }

        #endregion

        #region writeExcel

        public void GeneratFiles()
        {
            writeExcelAtmFile(this.AtmTable);
            writePrintExcelAtmFile(this.PrintAtmTable);
            this.writeCashFile(this.CashTable);
        }

        public string writeExcelAtmFile(DataTable dt)
        {

            IWorkbook wb = null;
            string ext = Path.GetExtension(this.FilePath);
            string fullname = Path.GetFullPath(this.FilePath);
            string name = fullname.Substring(fullname.Length - ext.Length);
            string CreatedFilePath = fullname + "ATM" + ext;
            try
            {


                using (FileStream stream = new FileStream(CreatedFilePath, FileMode.Create, FileAccess.Write))
                {
                    //Creat new Workbook Depend on file type
                    if (this.hssWB.GetType() == typeof(XSSFWorkbook))
                    {
                        wb = new XSSFWorkbook();
                    }
                    else
                    {
                        wb = new HSSFWorkbook();
                    }


                    ISheet sheet = wb.CreateSheet("Sheet1");
                    ICellStyle headerStyle = this.style.CreateHeaderCodeStyle(wb);
                    ICellStyle contentStyle = this.style.CreateContentCellStyle(wb);
                    sheet.IsRightToLeft = true;
                    ICreationHelper cH = wb.GetCreationHelper();

                    IRow coderow = sheet.CreateRow(0);
                    IRow namerow = sheet.CreateRow(1);
                    int i = 0;
                    foreach (DataColumn col in dt.Columns)
                    {

                        string[] colItems = new string[2];
                        colItems = col.ColumnName.Split('-');

                        coderow.CreateCell(i).SetCellValue(int.Parse(colItems[0]));
                        coderow.GetCell(i).CellStyle = headerStyle;

                        namerow.CreateCell(i).SetCellValue(colItems[1]);
                        namerow.GetCell(i).CellStyle = headerStyle;
                        sheet.AutoSizeColumn(i);
                        i++;
                    }
                    int currentRowIndex = 2;
                    foreach (DataRow row in dt.Rows)
                    {
                        IRow currentRow = sheet.CreateRow(currentRowIndex);
                        int currentCell = 0;
                        foreach (var cell in row.ItemArray)
                        {
                            double cellVal = 0;
                            bool convert = double.TryParse(cell.ToString(), out cellVal);
                            ICell myCell = currentRow.CreateCell(currentCell);
                            if (convert)
                            {

                                myCell.SetCellType(CellType.Numeric);
                                myCell.SetCellValue(cellVal);
                            }
                            else
                            {
                                if (cell != null)
                                {
                                    myCell.SetCellValue(cell.ToString());
                                }
                                else
                                {

                                    myCell.SetCellValue(string.Empty);
                                }
                            }
                            myCell.CellStyle = contentStyle;
                            currentCell++;
                        }

                        currentRowIndex++;

                    }
                    try
                    {
                        int i2 = 0;
                        foreach (var col in dt.Columns)
                        {
                           sheet.SetColumnWidth(i2,4000);
                            i2++;
                        }


                        wb.Write(stream);
                    }
                    catch (Exception ex)
                    {

                        MessageBox.Show(ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


            Dictionary<int, string> headerCols = new Dictionary<int, string>();
            foreach (DataColumn col in dt.Columns)
            {
                if (col.ColumnName.Contains("-"))
                {
                    string[] headeritems = new string[2];
                    headeritems = col.ColumnName.Split('-');
                    try
                    {
                        headerCols.Add(int.Parse(headeritems[0]), headeritems[1]);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message + headeritems[0]);

                    }

                }
            }

            return null;
        }
        public string writePrintExcelAtmFile(DataTable dt)
        {

            IWorkbook wb = null;
            string ext = Path.GetExtension(this.FilePath);
            string fullname = Path.GetFullPath(this.FilePath);
            string name = fullname.Substring(fullname.Length - ext.Length);
            string CreatedFilePath = fullname + "ATMPrint" + ext;
            try
            {


                using (FileStream stream = new FileStream(CreatedFilePath, FileMode.Create, FileAccess.Write))
                {
                    //Creat new Workbook Depend on file type
                    if (this.hssWB.GetType() == typeof(XSSFWorkbook))
                    {
                        wb = new XSSFWorkbook();
                    }
                    else
                    {
                        wb = new HSSFWorkbook();
                    }


                    ISheet sheet = wb.CreateSheet("Sheet1");
                    ICellStyle contentStyle = this.style.CreateContentCellStyle(wb);
                    sheet.IsRightToLeft = true;
                    ICreationHelper cH = wb.GetCreationHelper();

                    IRow namerow = sheet.CreateRow(0);

                    int i = 0;
                    foreach (DataColumn col in dt.Columns)
                    {

                        namerow.CreateCell(i).SetCellValue(col.ColumnName);
                        namerow.GetCell(i).CellStyle = contentStyle;
                        i++;
                    }
                    int currentRowIndex = 1;
                    foreach (DataRow row in dt.Rows)
                    {
                        IRow currentRow = sheet.CreateRow(currentRowIndex);
                        currentRow.CreateCell(0);
                        currentRow.GetCell(0).SetCellType(CellType.String);
                        currentRow.GetCell(0).SetCellValue(row[0].ToString());
                        currentRow.GetCell(0).CellStyle = contentStyle;



                        currentRow.CreateCell(1);
                        currentRow.GetCell(1).SetCellType(CellType.String);
                        currentRow.GetCell(1).SetCellValue(row[1].ToString());
                        currentRow.GetCell(1).CellStyle = contentStyle;



                        currentRow.CreateCell(2);
                        currentRow.GetCell(2).SetCellType(CellType.String);
                        currentRow.GetCell(2).SetCellValue(row[2].ToString());
                        currentRow.GetCell(2).CellStyle = contentStyle;



                        currentRow.CreateCell(3);
                        currentRow.GetCell(3).SetCellType(CellType.String);
                        currentRow.GetCell(3).SetCellValue(row[3].ToString());
                        currentRow.GetCell(3).CellStyle = contentStyle;



                        currentRow.CreateCell(4);
                        currentRow.GetCell(4).SetCellType(CellType.String);
                        currentRow.GetCell(4).SetCellValue(row[4].ToString());
                        currentRow.GetCell(4).CellStyle = contentStyle;




                        currentRow.CreateCell(5);
                        currentRow.GetCell(5).SetCellType(CellType.String);
                        currentRow.GetCell(5).SetCellValue(row[5].ToString());
                        currentRow.GetCell(5).CellStyle = contentStyle;


                        currentRow.CreateCell(6);
                        currentRow.GetCell(6).SetCellType(CellType.Numeric);
                        currentRow.GetCell(6).SetCellValue(double.Parse( row[6].ToString()));
                        currentRow.GetCell(6).CellStyle = contentStyle;

                        currentRowIndex++;

                    }
                    try
                    {

                        sheet.SetColumnWidth(0,5000);
                        sheet.SetColumnWidth(1, 5000);
                        sheet.SetColumnWidth(2, 2000);
                        sheet.SetColumnWidth(3, 4000);
                        sheet.SetColumnWidth(4, 4000);
                        sheet.SetColumnWidth(5, 9000);
                        sheet.SetColumnWidth(6, 3000);

                        wb.Write(stream);
                    }
                    catch (Exception ex)
                    {

                        MessageBox.Show(ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


            Dictionary<int, string> headerCols = new Dictionary<int, string>();
            foreach (DataColumn col in dt.Columns)
            {
                if (col.ColumnName.Contains("-"))
                {
                    string[] headeritems = new string[2];
                    headeritems = col.ColumnName.Split('-');
                    try
                    {
                        headerCols.Add(int.Parse(headeritems[0]), headeritems[1]);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message + headeritems[0]);

                    }

                }
            }


            return "";
        }

        public string writeCashFile(DataTable dt)
        {

            IWorkbook wb = null;
            string ext = Path.GetExtension(this.FilePath);
            string fullname = Path.GetFullPath(this.FilePath);
            string name = fullname.Substring(fullname.Length - ext.Length);
            string CreatedFilePath = fullname + "Cash" + ext;
            try
            {


                using (FileStream stream = new FileStream(CreatedFilePath, FileMode.Create, FileAccess.Write))
                {
                    //Creat new Workbook Depend on file type
                    if (this.hssWB.GetType() == typeof(XSSFWorkbook))
                    {
                        wb = new XSSFWorkbook();
                    }
                    else
                    {
                        wb = new HSSFWorkbook();
                    }


                    ISheet sheet = wb.CreateSheet("Sheet1");
                    ICellStyle contentStyle = this.style.CreateContentCellStyle(wb);
                    ICellStyle titleStyle = this.style.TitleCellsStyle(wb);
                    sheet.IsRightToLeft = true;
                    ICreationHelper cH = wb.GetCreationHelper();
                    sheet.CreateRow(0);
                    sheet.CreateRow(1);
                    sheet.CreateRow(2);
                    sheet.CreateRow(4);
                    sheet.GetRow(0).CreateCell(0);
                    sheet.GetRow(0).Cells[0].SetCellValue("جامعة الأسكندرية");
                    sheet.GetRow(0).Cells[0].CellStyle = titleStyle;

                    sheet.GetRow(1).CreateCell(0);
                    sheet.GetRow(1).Cells[0].SetCellValue("كلية الطب");
                    sheet.GetRow(1).Cells[0].CellStyle = titleStyle;


                    sheet.GetRow(2).CreateCell(0);
                    sheet.GetRow(2).Cells[0].SetCellValue("الوحدة الحسابية");
                    sheet.GetRow(2).Cells[0].CellStyle = titleStyle;

                    CellRangeAddress cra = new CellRangeAddress(4,5,0,2);
                    sheet.AddMergedRegion(cra);
                    sheet.GetRow(4).CreateCell(0);
                  
                    sheet.GetRow(4).Cells[0].SetCellValue("كشف صرف شيكات للسادة           عن استمارة             شهر         ");
                    sheet.GetRow(4).Cells[0].CellStyle = this.style.CreateNameCodeStyle(wb);
                    IRow namerow = sheet.CreateRow(6);

                    int i = 0;
                    foreach (DataColumn col in dt.Columns)
                    {

                        namerow.CreateCell(i).SetCellValue(col.ColumnName);
                        namerow.Cells[i].CellStyle = contentStyle;
                        i++;
                    }
                    int currentRowIndex = 7;
                    foreach (DataRow row in dt.Rows)
                    {
                        IRow currentRow = sheet.CreateRow(currentRowIndex);
                        currentRow.CreateCell(0);
                        currentRow.CreateCell(1);
                        currentRow.Cells[1].SetCellType(CellType.Numeric);
                        currentRow.CreateCell(2);
                        currentRow.Cells[0].SetCellValue(row[0].ToString());
                        currentRow.Cells[0].CellStyle = contentStyle;

                        currentRow.Cells[1].SetCellValue(double.Parse(row[1].ToString()));
                        currentRow.Cells[1].CellStyle = contentStyle;


                        currentRow.Cells[2].SetCellValue(row[2].ToString());
                        currentRow.Cells[2].CellStyle = contentStyle;

                        currentRowIndex++;

                    }

                    IRow totalRow = sheet.CreateRow(currentRowIndex);
                    totalRow.CreateCell(0);
                    totalRow.Cells[0].CellStyle = contentStyle;
                    totalRow.CreateCell(1);
                    totalRow.Cells[1].CellStyle = contentStyle;
                    totalRow.Cells[0].SetCellValue(" الأجمالى ");
                    totalRow.CreateCell(2);
                    totalRow.Cells[2].CellStyle = contentStyle;
                    if (currentRowIndex > 8)
                    {

                        totalRow.Cells[1].SetCellFormula("SUM(B8:B" + (currentRowIndex) + ")");
                    }
                    else
                    {
                        totalRow.Cells[1].SetCellValue(string.Empty);
                    }
                    try
                    {
                        sheet.SetColumnWidth(0, 8000);
                        sheet.SetColumnWidth(1, 4000);
                        sheet.SetColumnWidth(2, 8000);
                        wb.Write(stream);
                    }
                    catch (Exception ex)
                    {

                        MessageBox.Show(ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


            Dictionary<int, string> headerCols = new Dictionary<int, string>();
            foreach (DataColumn col in dt.Columns)
            {
                if (col.ColumnName.Contains("-"))
                {
                    string[] headeritems = new string[2];
                    headeritems = col.ColumnName.Split('-');
                    try
                    {
                        headerCols.Add(int.Parse(headeritems[0]), headeritems[1]);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message + headeritems[0]);

                    }

                }
            }



            return "";
        }
        #endregion
    }

    public class HeaderList
    {
        public List<HeaderContent> headerContent { get; set; }

        public HeaderList()
        {
            this.headerContent = new List<HeaderContent>();
        }
    }

    public class HeaderContent
    {
        public int HeaderCode { get; set; }

        public string HeaderName { get; set; }

        public int ColIndex { get; set; }

        public string HeaderType { get; set; }
    }

}
