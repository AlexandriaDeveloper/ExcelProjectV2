using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProjectV2.NPOI
{
    using System.Collections;
    using System.Data;
    using System.Drawing;
    using System.IO;
    using System.Windows.Forms;

    using ExcelProjectV2.Model;

    using global::NPOI.HSSF.UserModel;
    using global::NPOI.HSSF.Util;
    using global::NPOI.SS.UserModel;
    using global::NPOI.XSSF.UserModel;

    using BorderStyle = global::NPOI.SS.UserModel.BorderStyle;
    using HorizontalAlignment = global::NPOI.SS.UserModel.HorizontalAlignment;

    public class NPOIHelper : IDisposable
    {
        #region Local Variables

        private string _FilePath;

        private IWorkbook hssWB = null;

        private ISheet sheet = null;

        private ICell cell = null;

        private IRow row = null;

        private IFormulaEvaluator helper = null;

        private DataFormatter df = null;

        private DataTable dt = null;

        private int startrowindex;

        private NpoiStyle cellStyle = null;

        DataTable EpaymentTable = null;

        DataTable PrintTable = null;

        DataTable CashTable = null;

        #endregion

        #region  Constracture

        public NPOIHelper(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("من فضلك أختار ملف ");
            }
            //Announce File Path
            _FilePath = filePath;
            this.cellStyle = new NpoiStyle();
            string ext = Path.GetExtension(filePath).ToLower();
            if (ext != ".xls" && ext != ".xlsx")
            {
                MessageBox.Show("عفوا لا يمكن التعامل مع صيغة هذا الملف ");
            }
            this._FilePath = filePath;

            //Create WorkBook Depend on Excel File Type xls or xlsx
            try
            {
                using (Stream file = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
                {
                    if (ext == ".xls")
                    {
                        hssWB = new HSSFWorkbook(file);
                        HSSFFormulaEvaluator.EvaluateAllFormulaCells(hssWB);

                        helper = (HSSFFormulaEvaluator)hssWB.GetCreationHelper().CreateFormulaEvaluator();
                        this.helper.EvaluateAll();
                        file.Close();
                    }
                    else if (ext == ".xlsx")
                    {
                        this.hssWB = new XSSFWorkbook(file);
                        XSSFFormulaEvaluator.EvaluateAllFormulaCells(hssWB);

                        helper = (XSSFFormulaEvaluator)hssWB.GetCreationHelper().CreateFormulaEvaluator();
                        this.helper.EvaluateAll();
                        file.Close();

                    }

                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }


        #endregion




        internal void WriteExcelFile(DataTable dt)
        {
            this.EpaymentTable = new DataTable();
            this.CashTable = new DataTable();

            Hashtable ColIndex = new Hashtable();
            foreach (DataColumn dataColumn in dt.Columns)
            {

                EpaymentTable.Columns.Add(dataColumn.ColumnName);
                if (dataColumn.ColumnName.StartsWith("3333") || //الصافى
                    dataColumn.ColumnName.StartsWith("3001") || //الأسم
                    dataColumn.ColumnName.StartsWith("3003")) //الكود
                {
                    ColIndex.Add(dataColumn.ColumnName, dataColumn.Ordinal);
                }



            }
            this.CashTable.Columns.Add(new DataColumn("3001-الأسم"));
            this.CashTable.Columns.Add(new DataColumn("3333-الصافى"));
            this.CashTable.Columns.Add(new DataColumn("0000-التوقيع"));


            this.PrintTable = new DataTable();
            this.PrintTable.Columns.Add(new DataColumn("الرقم القومى"));
            this.PrintTable.Columns.Add(new DataColumn("نوع المدفوعه"));
            this.PrintTable.Columns.Add(new DataColumn("القطاع"));
            this.PrintTable.Columns.Add(new DataColumn("الإدارة"));
            this.PrintTable.Columns.Add(new DataColumn("رقم الموظف بجهته الأصلية"));
            this.PrintTable.Columns.Add(new DataColumn("الاسم"));
            this.PrintTable.Columns.Add(new DataColumn("المرتب"));

            foreach (DataRow dataRow in dt.Rows)
            {

                int i = 0;
                if (dataRow.ItemArray[(int)ColIndex["3003-الكود"]] == ""
                    || dataRow.ItemArray[(int)ColIndex["3003-الكود"]].ToString() == "0")

                {
                    WriteCashFile(dataRow);
                }
                else
                {
                    WriteAtmFile(dataRow);
                    WritePrintFile(dataRow);

                }



            }
            var filenamepath = Path.GetFullPath(_FilePath); //+ "Atm" + Path.GetExtension(this._FilePath);
            filenamepath = Path.ChangeExtension(filenamepath, null);


            if (this.EpaymentTable.Rows.Count > 0)
            {
                string atmFileNmae = filenamepath + "Atm" + Path.GetExtension(this._FilePath);
                string PrintatmFileNmae = filenamepath + "PrintAtm" + Path.GetExtension(this._FilePath);
                WriteExcelFile(this.EpaymentTable, atmFileNmae);
                WritePrintExcelFile(this.PrintTable, PrintatmFileNmae);
            }

            if (this.CashTable.Rows.Count > 0)
            {

                string cashFileNmae = filenamepath + "cash" + Path.GetExtension(this._FilePath);
                IWorkbook wb = WriteExcelFile(this.CashTable, cashFileNmae, 4);
                //    GeneratLogo(wb, cashFileNmae, "Sheet1");
                //   GeneratSumTotal(wb,cashFileNmae, "Sheet1");

            }
            MessageBox.Show("تم بنجاح");

        }

        private void WritePrintFile(DataRow dr)
        {
            int CodeIndex = dr.Table.Columns["3003-الكود"].Ordinal;
            int NameIndex = dr.Table.Columns["3001-الأسم"].Ordinal;
            int NetIndex = dr.Table.Columns["3333-الصافى"].Ordinal;



            DataRow drow = this.PrintTable.NewRow();

            int i = 0;
            // string EmpName = string.Empty;


            var Emp = this.context.Employees.Find(int.Parse(dr[CodeIndex].ToString()));
            drow[6] = dr[NetIndex];
            drow[5] = Emp.Name;
            drow[4] = Emp.Code;
            drow[3] = Emp.PositionName;
            drow[2] = string.Empty;
            drow[1] = "2-اخرى بطاقات حكومية";
            drow[0] = Emp.NationalId;





            this.PrintTable.Rows.Add(drow);

        }

        private void GeneratLogo(ISheet sheet)
        {
            ISheet currentSheet = sheet;
            currentSheet.ShiftRows(0, 2, 4);
            //IRow currentRow = currentSheet.CreateRow(0);
            //ICell currentCell = null;
            //currentCell = currentRow.CreateCell(0, CellType.String);
            //currentCell.SetCellValue("جامعة الأسكندرية");
            //currentRow = currentSheet.CreateRow(1);
            //currentCell = currentRow.CreateCell(0, CellType.String);
            //currentCell.SetCellValue("كلية الطب");
            //currentRow = currentSheet.CreateRow(2);
            //currentCell = currentRow.CreateCell(0, CellType.String);
            //currentCell.SetCellValue("الوحدة الحسابية ");
            //currentRow = currentSheet.CreateRow(3);
            //currentCell = currentRow.CreateCell(2, CellType.String);
            //currentCell.SetCellValue("كشف  مندوب صرف مستحقات السادة /                عن شهر  ");

        }

        private void GeneratSumTotal(IWorkbook wb, string FilePath, string sheetname)
        {
            ISheet currentSheet = wb.GetSheet(sheetname);
            GeneratLogo(currentSheet);

            //IRow currentRow = currentSheet.CreateRow(currentSheet.LastRowNum + 1);
            //ICell currentCell = null;
            //currentCell = currentRow.CreateCell(0, CellType.String);
            //currentCell.SetCellValue("الأجمالى");
            //currentCell.CellStyle = CreateHeaderCodeStyle(wb);
            //currentCell = currentRow.CreateCell(1,CellType.Formula);
            //currentCell.SetCellFormula(string.Format("Sum(b7:b{0})",currentSheet.LastRowNum));
            //currentCell.CellStyle = CreateHeaderCodeStyle(wb);

            using (FileStream stream = new FileStream(FilePath, FileMode.Open, FileAccess.ReadWrite))
            {
                wb.Write(stream);
                wb.Close();

            }


        }

        ERPEntities context = new ERPEntities();

        private bool WriteAtmFile(DataRow dr)
        {

            int CodeIndex = dr.Table.Columns["3003-الكود"].Ordinal;
            int NameIndex = dr.Table.Columns["3001-الأسم"].Ordinal;
            DataRow drow = this.EpaymentTable.NewRow();
            int i = 0;
            string EmpName = string.Empty;
            foreach (var item in dr.ItemArray)
            {
                if (CodeIndex == i)
                {
                    EmpName = this.context.Employees.Find(int.Parse(item.ToString())).Name;
                    drow[NameIndex] = EmpName;
                    drow[i] = item;
                }
                else
                {
                    drow[i] = item;
                }


                i++;
            }
            this.EpaymentTable.Rows.Add(drow);



            return true;
        }

        private IWorkbook WriteExcelFile(DataTable dt, string filePath, int statrow = 0)
        {
            IWorkbook wb = null;
            try
            {


                using (FileStream stream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
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
                    sheet.IsRightToLeft = true;
                    ICreationHelper cH = wb.GetCreationHelper();

                    IRow coderow = sheet.CreateRow(statrow);
                    IRow namerow = sheet.CreateRow(statrow + 1);

                    var headerCodeStyle = cellStyle.CreateHeaderCodeStyle(wb);
                    var headernameStyle = cellStyle.CreateNameCodeStyle(wb);
                    var contentStyle = cellStyle.CreateContentCellStyle(wb);

                    if (statrow == 4)
                    {
                        sheet.DefaultColumnWidth = 30;
                        IRow TitleRow1 = sheet.CreateRow(0);
                        ICell CellRow1 = TitleRow1.CreateCell(0);
                        CellRow1.SetCellValue("جامعة الأسكندرية");
                        CellRow1.CellStyle = cellStyle.TitleCellsStyle(wb);
                        TitleRow1 = sheet.CreateRow(1);
                        CellRow1 = TitleRow1.CreateCell(0);
                        CellRow1.CellStyle = cellStyle.TitleCellsStyle(wb);
                        CellRow1.SetCellValue("كلية الطب");
                        CellRow1.CellStyle = cellStyle.TitleCellsStyle(wb);
                        TitleRow1 = sheet.CreateRow(2);
                        CellRow1 = TitleRow1.CreateCell(0);
                        CellRow1.SetCellValue("الوحدة الحسابية");
                        CellRow1.CellStyle = cellStyle.TitleCellsStyle(wb);

                    }



                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        string[] ColumnName = dt.Columns[j].ToString().Split('-');
                        int headerCode = int.Parse(ColumnName[0]);
                        string headerName = ColumnName[1];
                        ICell codeCellcell = coderow.CreateCell(j);
                        codeCellcell.CellStyle = headerCodeStyle;


                        ICell Namecell = namerow.CreateCell(j);
                        Namecell.CellStyle = headernameStyle;
                        codeCellcell.SetCellValue(headerCode);
                        Namecell.SetCellValue(headerName);

                        //  co.SetCellValue(cH.CreateRichTextString(dt.Columns[j].ToString()));
                    }


                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        IRow row = sheet.CreateRow(i + statrow + 2);
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            row.Height = 400;
                            ICell cell = row.CreateCell(j);
                            cell.CellStyle = contentStyle;
                            double amount;
                            bool cellint = double.TryParse(dt.Rows[i].ItemArray[j].ToString(), out amount);
                            if (cellint)
                            {
                                cell.SetCellValue(amount);
                            }

                            else
                            {
                                cell.SetCellValue(dt.Rows[i].ItemArray[j].ToString());
                            }
                            // cell.SetCellValue(cH.CreateRichTextString(dt.Rows[i].ItemArray[j].ToString()));
                        }
                    }
                    if (statrow == 4)
                    {
                        IRow SumRow = sheet.CreateRow(sheet.LastRowNum + 1);
                        ICell SumCell = SumRow.CreateCell(1);
                        SumCell.SetCellType(CellType.Formula);
                        SumCell.SetCellFormula("Sum(b7:b" + sheet.LastRowNum + ")");
                        SumCell.CellStyle = cellStyle.CreateContentCellStyle(wb);
                        SumCell = SumRow.CreateCell(0);

                        SumCell.SetCellValue("الأجمالى");
                        SumCell.CellStyle = cellStyle.CreateContentCellStyle(wb);

                        SumCell = SumRow.CreateCell(2);
                        SumCell.SetCellValue(string.Empty);
                        SumCell.CellStyle = cellStyle.CreateContentCellStyle(wb);

                    }


                    try
                    {
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
            return wb;

        }

        private IWorkbook WritePrintExcelFile(DataTable dt, string filePath)
        {
            IWorkbook wb = null;
            try
            {


                using (FileStream stream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
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
                    sheet.IsRightToLeft = true;
                    sheet.SetColumnWidth(6, 20 * 256);
                    sheet.SetColumnWidth(5, 40 * 256);
                    sheet.SetColumnWidth(4, 20 * 256);
                    sheet.SetColumnWidth(3, 20 * 256);
                    sheet.SetColumnWidth(2, 10 * 256);
                    sheet.SetColumnWidth(1, 20 * 256);
                    sheet.SetColumnWidth(0, 30 * 256);

                    ICreationHelper cH = wb.GetCreationHelper();

                    IRow headerRow = sheet.CreateRow(0);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        string ColumnName = dt.Columns[j].ToString();
                        ICell codeCellcell = headerRow.CreateCell(j);
                        codeCellcell.SetCellValue(ColumnName);
                        codeCellcell.CellStyle = cellStyle.CreateNameCodeStyle(wb);
                    }
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        IRow row = sheet.CreateRow(i + 1);
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            row.Height = 400;
                            ICell cell = row.CreateCell(j);
                            double amount = 0;
                            bool cellint = double.TryParse(dt.Rows[i].ItemArray[j].ToString(), out amount);
                            if (cellint && j != 0)
                            {
                                cell.SetCellType(CellType.Numeric);
                                cell.SetCellValue(amount);
                                cell.CellStyle = cellStyle.CreateContentCellStyle(wb);
                            }
                            else
                            {
                                cell.SetCellType(CellType.String);
                                cell.SetCellValue(dt.Rows[i].ItemArray[j].ToString());
                                cell.CellStyle = cellStyle.CreateContentCellStyle(wb);
                            }

                        }
                    }
                    try
                    {
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
            return wb;

        }

        private bool WriteCashFile(DataRow dr)
        {
            DataRow drow = this.CashTable.NewRow();

            drow[0] = dr["3001-الأسم"];
            drow[1] = dr["3333-الصافى"];
            drow[2] = string.Empty;
            this.CashTable.Rows.Add(drow);


            return true;
        }

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



        public ExcelHeader GetExcelHeader(string sheetname, int rownum = 9)
        {
            startrowindex = rownum;
            sheet = hssWB.GetSheet(sheetname);
            this.row = this.sheet.GetRow(rownum);
            if (this.row == null)

            {
                MessageBox.Show("تأكد من اختيار السطر الصحيح");
                return null;
            }
            var rowvalues = this.sheet.GetRow(rownum + 1);
            df = new DataFormatter();
            ExcelHeader Header = new ExcelHeader();

            foreach (var cell in this.row)
            {
                string cellValu2e = df.FormatCellValue(cell);
                int cellValue;
                bool cellNumeric = int.TryParse(cellValu2e, out cellValue);

                ExcelHeaderContent headerContent = new ExcelHeaderContent();
                if (!string.IsNullOrEmpty(cellValue.ToString()) && cellNumeric)
                {


                    //if (this.cell.CellType == CellType.String)
                    //{
                    //    headerContent.HeaderCode = cell.StringCellValue;
                    //}
                    // if (this.cell.CellType == CellType.Numeric)
                    //{

                    var acc = this.context.Accounts.FirstOrDefault(x => x.Id == cellValue);
                    if (acc == null & !string.IsNullOrEmpty(cellValue.ToString()))
                    {

                        ///TODO Remove Hashing 
                        //MessageBox.Show(
                        //    string.Format(
                        //        "كود الحساب {0} غير موجود{1}",
                        //        cellValue.ToString(),
                        //        this.sheet.GetRow(rownum + 1).GetCell(cell.ColumnIndex)));
                    }


                    headerContent.HeaderCode = cellValue.ToString();

                    if (cellValue.ToString().StartsWith("1"))
                    {
                        headerContent.AccouuntType = "Credit";
                    }
                    else if (cellValue.ToString().StartsWith("2"))
                    {
                        headerContent.AccouuntType = "Debit";
                    }
                    else if (cellValue.ToString().StartsWith("3"))
                    {
                        headerContent.AccouuntType = "Def";
                    }
                    //  }
                    if (rowvalues.GetCell(cell.ColumnIndex) != null)
                    {
                        var headerTitle = this.sheet.GetRow(cell.RowIndex + 1).GetCell(cell.ColumnIndex).ToString();
                        headerContent.HeaderName = headerTitle;
                        headerContent.HeaderCellIndex = cell.ColumnIndex;
                    }



                    Header.ExcelHeaderContents.Add(headerContent);

                }
            }
            return Header;
        }

 

        public DataTable GetExceRows(List<string> sheetsname, ExcelHeader header)
        {

            int NameIndex = 0;
            int NetIndex = 0;
            int DueIndex = 0;
            int OutIndex = 0;
            //Creat Table Columns And Defin General Accounts Code
            dt = new DataTable();
            df = new DataFormatter();
            try
            {
                this.GenerateTableHeader(header, dt, ref NameIndex, ref DueIndex, ref OutIndex, ref NetIndex);
                foreach (var sheetname in sheetsname)
                {
                    
                    this.sheet = this.hssWB.GetSheet(sheetname);
                    if (this.sheet.GetRow(this.startrowindex) != null)
                    {
                        CheckHeaderCodeIndex(header, this.sheet.GetRow(this.startrowindex));
                    }
                    int startrow = this.startrowindex + 2;
                    //IRow TotalRow = this.sheet.GetRow(this.sheet.LastRowNum);
                    int lastRow = 0;
                    for (int i = startrow; i < this.sheet.LastRowNum; i++)
                    {
                        if (this.sheet.GetRow(i) != null &&
                            this.sheet.GetRow(i).GetCell(NameIndex).StringCellValue == "الأجمالى" && NameIndex > 0)
                        {
                            lastRow = i;
                            break;

                        }
                    }
                    if (lastRow == 0)
                    {
                        lastRow = this.sheet.LastRowNum;
                    }
                    IRow TotalRow = this.sheet.GetRow(lastRow);
                    decimal totalNet = 0;
                    decimal totalDue = 0;
                    decimal totalOut = 0;
                    for (int i = startrow; i < lastRow; i++)
                    {
                        DataRow dr = dt.NewRow();
                        int col = 0;
                        IRow CurrentRow = this.sheet.GetRow(startrow);
                        decimal rownet = 0;
                        decimal rowDue = 0;
                        decimal rowOut = 0;
                        if (CurrentRow != null )
                            foreach (var colIndex in header.ExcelHeaderContents)
                            {
                                var currentCell = CurrentRow.GetCell(colIndex.HeaderCellIndex);
                                if (currentCell == null)
                                {
                                    currentCell = CurrentRow.CreateCell(colIndex.HeaderCellIndex);
                                }


                                if (currentCell.CellType == CellType.Formula)
                                {
                                    //   this.helper.EvaluateFormulaCell(currentCell);


                                    //  this.helper.EvaluateInCell(currentCell);

                                    if (this.hssWB.GetType() == typeof(HSSFWorkbook))
                                    {
                                        this.helper = new HSSFFormulaEvaluator(this.hssWB);

                                    }
                                    else
                                    {
                                        this.helper = new XSSFFormulaEvaluator(this.hssWB);
                                    }
                                    this.helper.EvaluateInCell(currentCell);
                                }

                                string currentCellValue = df.FormatCellValue(currentCell);
                                if (currentCellValue.StartsWith("#"))
                                {
                                    currentCellValue = string.Empty;
                                }
                                if (string.IsNullOrEmpty(currentCellValue))
                                {
                                    dr[col] = (string.Empty);
                                }
                                else
                                {
                                    dr[col] = (currentCellValue);
                                    var t =
                                        header.ExcelHeaderContents.SingleOrDefault(
                                            x => x.HeaderCellIndex == currentCell.ColumnIndex);
                                    if (t != null)
                                    {

                                        rowDue += ReturnValueByAccountType(t, "Credit", currentCellValue);


                                        rowOut += ReturnValueByAccountType(t, "Debit", currentCellValue);



                                        CheckRowValuesEqulaity(
                                            t,
                                            "3111",
                                            rowDue,
                                            currentCellValue,
                                            CurrentRow.GetCell(NameIndex).StringCellValue);


                                        CheckRowValuesEqulaity(
                                            t,
                                            "3222",
                                            rowOut,
                                            currentCellValue,
                                            CurrentRow.GetCell(NameIndex).StringCellValue);
                                        rownet = rowDue - rowOut;
                                        CheckRowValuesEqulaity(
                                            t,
                                            "3333",
                                            rownet,
                                            currentCellValue,
                                            CurrentRow.GetCell(NameIndex).StringCellValue);
                                    }
                                }

                                //dr[col] = currentCellValue;

                                if (colIndex.HeaderCellIndex == NetIndex && !string.IsNullOrEmpty(currentCellValue))
                                {

                                    totalNet += decimal.Parse(currentCellValue);
                                }
                                if (colIndex.HeaderCellIndex == DueIndex && !string.IsNullOrEmpty(currentCellValue))
                                {

                                    totalDue += decimal.Parse(currentCellValue);
                                }
                                if (colIndex.HeaderCellIndex == OutIndex && !string.IsNullOrEmpty(currentCellValue))
                                {

                                    totalOut += decimal.Parse(currentCellValue);
                                }
                                col++;


                            }
                        dt.Rows.Add(dr);
                        startrow++;
                    }


                    if (totalDue != (decimal)TotalRow.GetCell(DueIndex).NumericCellValue)
                    {
                        MessageBox.Show(
                            string.Format(
                                "Sheet Name {0} , Due = {1} and SumResult = {2}",
                                sheetname,
                                TotalRow.GetCell(DueIndex).NumericCellValue.ToString(),
                                totalDue));
                    }
                    if (totalOut != (decimal)TotalRow.GetCell(OutIndex).NumericCellValue)
                    {
                        MessageBox.Show(
                            string.Format(
                                "Sheet Name {0} , Out = {1} and SumResult = {2}",
                                sheetname,
                                TotalRow.GetCell(OutIndex).NumericCellValue.ToString(),
                                totalOut));
                    }
                    if (totalNet != (decimal)TotalRow.GetCell(NetIndex).NumericCellValue)
                    {
                        MessageBox.Show(
                            string.Format(
                                "Sheet Name {0} , Net = {1} and SumResult = {2}",
                                sheetname,
                                TotalRow.GetCell(NetIndex).NumericCellValue.ToString(),
                                totalNet));
                    }


                }

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            return dt;

        }

        private decimal ReturnValueByAccountType(ExcelHeaderContent t, string AccountType, string CurrentCellValue)
        {

            if (AccountType == "Def")
            {
                if (t.AccouuntType == AccountType && t.HeaderCode == "3333")
                {
                }
            }

            if (t.AccouuntType == AccountType)

                return decimal.Parse(CurrentCellValue);
            return 0;
        }

        private void CheckRowValuesEqulaity(ExcelHeaderContent t, string AccountCode, decimal originalValue, string sumValue, string Name)
        {
            if (t.AccouuntType == "Def" && t.HeaderCode == AccountCode)
            {
                if (originalValue != decimal.Parse(sumValue))
                {
                    MessageBox.Show(
                        String.Format(
                            "{3} Should be {0} instead of {1} for emp {2}",
                            originalValue,
                          sumValue,
                            Name, t.HeaderName));
                }
            }
        }

        private void CheckHeaderCodeIndex(ExcelHeader header, IRow getRow)
        {
            DataFormatter df = new DataFormatter();
            foreach (var cell in getRow)
            {
                string cellValue = df.FormatCellValue(cell);

                if (string.IsNullOrEmpty(cellValue))
                {

                    ExcelHeaderContent hc = new ExcelHeaderContent();
                    hc =
                        header.ExcelHeaderContents.Where(x => x.HeaderCode == cellValue)
                            .FirstOrDefault();
                    if (hc != null)
                    {
                        if (hc.HeaderCellIndex != cell.ColumnIndex)
                        {
                            MessageBox.Show(
                                string.Format(
                                    "{0}  should be {1} in sheet {2}",
                                    hc.HeaderName.ToString(),
                                    cell.ColumnIndex.ToString(), getRow.Sheet.SheetName));
                        }
                    }
                }

            }
        }

        private void GenerateTableHeader(
            ExcelHeader header,
            DataTable dt,
           ref int NameIndex,
            ref int DueIndex,
            ref int OutIndex,
            ref int NetIndex)
        {
            if (header != null)
            {
                foreach (var colHeader in header.ExcelHeaderContents)
                {
                    dt.Columns.Add(colHeader.HeaderCode + "-" + colHeader.HeaderName);
                    if (colHeader.HeaderCode == "3001")
                    {
                        NameIndex = colHeader.HeaderCellIndex;
                    }
                    if (colHeader.HeaderCode == "3111")
                    {
                        DueIndex = colHeader.HeaderCellIndex;
                    }
                    if (colHeader.HeaderCode == "3222")
                    {
                        OutIndex = colHeader.HeaderCellIndex;
                    }
                    if (colHeader.HeaderCode == "3333")
                    {
                        NetIndex = colHeader.HeaderCellIndex;
                    }
                }
            }
            else
            {
                MessageBox.Show("تأكد من أختيار السط المناسب");
            }
        }

        //#region  Cell Styling
        //private ICellStyle CreateHeaderCodeStyle(IWorkbook wb)
        //{
        //    IFont FontCode = wb.CreateFont();
        //    FontCode.Boldweight = (short)FontBoldWeight.Bold;
        //    FontCode.FontName = "Andalus";
        //    FontCode.FontHeightInPoints = 8;
        //    ICellStyle boldStyle = wb.CreateCellStyle();
        //    // boldStyle.IsHidden = true;
        //    boldStyle.Alignment = HorizontalAlignment.Center;
        //    boldStyle.FillForegroundColor = 25;
        //    SetHeaderBorder(boldStyle, BorderStyle.Thick);
        //    boldStyle.SetFont(FontCode);
        //    boldStyle.FillPattern = FillPattern.SolidForeground;
        //    return boldStyle;

        //}
        //private ICellStyle CreateNameCodeStyle(IWorkbook wb)
        //{
        //    IFont FontCode = wb.CreateFont();
        //    FontCode.Boldweight = (short)FontBoldWeight.Bold;
        //    FontCode.FontName = "Arial";
        //    FontCode.FontHeightInPoints = 16;
        //    ICellStyle boldStyle = wb.CreateCellStyle();
        //    boldStyle.FillForegroundColor = 25;
        //    boldStyle.VerticalAlignment = VerticalAlignment.Center;
        //    boldStyle.Alignment = HorizontalAlignment.Center;
        //    SetHeaderBorder(boldStyle, BorderStyle.Thick);
        //    boldStyle.SetFont(FontCode);
        //    boldStyle.FillPattern = FillPattern.SolidForeground;
        //    return boldStyle;

        //}
        //private ICellStyle CreateContentCellStyle(IWorkbook wb)
        //{
        //    IFont FontCode = wb.CreateFont();
        //    FontCode.Boldweight = (short)FontBoldWeight.None;
        //    FontCode.FontName = "Arial";
        //    FontCode.FontHeightInPoints = 12;
        //    ICellStyle boldStyle = wb.CreateCellStyle();
        //    boldStyle.VerticalAlignment = VerticalAlignment.Center;
        //    boldStyle.Alignment = HorizontalAlignment.Center;
        //    SetHeaderBorder(boldStyle, BorderStyle.Thin);
        //    boldStyle.SetFont(FontCode);
        //    return boldStyle;

        //}
        //private ICellStyle TitleCellsStyle(IWorkbook wb)
        //{
        //    IFont FontCode = wb.CreateFont();
        //    FontCode.Boldweight = (short)FontBoldWeight.Bold;
        //    FontCode.FontName = "Andalus";
        //    FontCode.FontHeightInPoints = 13;
        //    ICellStyle boldStyle = wb.CreateCellStyle();
        //    // boldStyle.IsHidden = true;
        //    boldStyle.Alignment = HorizontalAlignment.Right;

        //    //boldStyle.FillForegroundColor = 25;
        //    //  SetHeaderBorder(boldStyle, BorderStyle.Thick);
        //    boldStyle.SetFont(FontCode);
        //    // boldStyle.FillPattern = FillPattern.SolidForeground;
        //    return boldStyle;

        //}

        //private void SetHeaderBorder(ICellStyle cell, BorderStyle borderStyle)
        //{
        //    cell.BorderBottom = borderStyle;
        //    cell.BorderRight = borderStyle;
        //    cell.BorderTop = borderStyle;
        //    cell.BorderLeft = borderStyle;

        //}

        //#endregion


        public void Dispose()
        {
            if (hssWB != null)
            {
                if (this.hssWB.GetType() == typeof(HSSFWorkbook))
                {
                    ((HSSFWorkbook)this.hssWB).Close();
                }
                if (this.hssWB.GetType() == typeof(XSSFWorkbook))
                {
                    ((XSSFWorkbook)this.hssWB).Close();


                }
            }
        }




    }



    public class ExcelHeader
    {
        public IList<ExcelHeaderContent> ExcelHeaderContents { get; set; }

        public ExcelHeader()
        {
            ExcelHeaderContents = new List<ExcelHeaderContent>();
        }
    }

    public class ExcelHeaderContent
    {
        public string HeaderCode { get; set; }

        public string HeaderName { get; set; }

        public string AccouuntType { get; set; }

        public int HeaderCellIndex { get; set; }
    }


}
