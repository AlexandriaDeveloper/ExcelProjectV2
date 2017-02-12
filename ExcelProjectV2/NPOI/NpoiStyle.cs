using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProjectV2.NPOI
{
    using global::NPOI.SS.UserModel;

    public class NpoiStyle
    {

        #region  Cell Styling
        public ICellStyle CreateHeaderCodeStyle(IWorkbook wb)
        {
            IFont FontCode = wb.CreateFont();
            FontCode.Boldweight = (short)FontBoldWeight.Bold;
            FontCode.FontName = "Andalus";
            FontCode.FontHeightInPoints = 8;
            ICellStyle boldStyle = wb.CreateCellStyle();
            // boldStyle.IsHidden = true;
            boldStyle.Alignment = HorizontalAlignment.Center;
            boldStyle.FillForegroundColor = 25;
            SetHeaderBorder(boldStyle, BorderStyle.Thick);
            boldStyle.SetFont(FontCode);
            boldStyle.FillPattern = FillPattern.SolidForeground;
            return boldStyle;

        }
        public ICellStyle CreateNameCodeStyle(IWorkbook wb)
        {
            IFont FontCode = wb.CreateFont();
            FontCode.Boldweight = (short)FontBoldWeight.Bold;
            FontCode.FontName = "Arial";
            FontCode.FontHeightInPoints = 12;
            ICellStyle boldStyle = wb.CreateCellStyle();
            //boldStyle.FillForegroundColor = 25;
            boldStyle.VerticalAlignment = VerticalAlignment.Center;
            boldStyle.Alignment = HorizontalAlignment.Center;
           // SetHeaderBorder(boldStyle, BorderStyle.Thick);
            boldStyle.SetFont(FontCode);
          //  boldStyle.FillPattern = FillPattern.SolidForeground;
            return boldStyle;

        }
        public ICellStyle CreateContentCellStyle(IWorkbook wb)
        {
            IFont FontCode = wb.CreateFont();
            FontCode.Boldweight = (short)FontBoldWeight.None;
            FontCode.FontName = "Arial";
            FontCode.FontHeightInPoints = 12;
            ICellStyle boldStyle = wb.CreateCellStyle();
            boldStyle.VerticalAlignment = VerticalAlignment.Center;
            boldStyle.Alignment = HorizontalAlignment.Center;
            SetHeaderBorder(boldStyle, BorderStyle.Thin);
            boldStyle.SetFont(FontCode);
            return boldStyle;

        }
        public ICellStyle TitleCellsStyle(IWorkbook wb)
        {
            IFont FontCode = wb.CreateFont();
            FontCode.Boldweight = (short)FontBoldWeight.Bold;
            FontCode.FontName = "Andalus";
            FontCode.FontHeightInPoints = 13;
            ICellStyle boldStyle = wb.CreateCellStyle();
            // boldStyle.IsHidden = true;
            boldStyle.Alignment = HorizontalAlignment.Right;

            //boldStyle.FillForegroundColor = 25;
            //  SetHeaderBorder(boldStyle, BorderStyle.Thick);
            boldStyle.SetFont(FontCode);
            // boldStyle.FillPattern = FillPattern.SolidForeground;
            return boldStyle;

        }

        private void SetHeaderBorder(ICellStyle cell, BorderStyle borderStyle)
        {
            cell.BorderBottom = borderStyle;
            cell.BorderRight = borderStyle;
            cell.BorderTop = borderStyle;
            cell.BorderLeft = borderStyle;

        }

        #endregion

    }
}
