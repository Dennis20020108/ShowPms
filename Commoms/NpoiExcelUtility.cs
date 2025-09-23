using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace ShowPms.Commoms
{
    public static class NpoiExcelUtility
    {
        public static ICell SetCell(ISheet sheet, int rowIndex, int colIndex, string value)
        {
            IRow targetRow = sheet.GetRow(rowIndex);
            if (targetRow == null)
            {
                targetRow = sheet.CreateRow(rowIndex);
            }

            ICell cell = targetRow.GetCell(colIndex);
            if (cell == null)
            {
                cell = targetRow.CreateCell(colIndex);
            }

            cell.SetCellValue(value);

            return cell;
        }

        public static ICell SetCellStyle(ISheet sheet, int rowIndex, int colIndex, ICellStyle style)
        {
            IRow targetRow = sheet.GetRow(rowIndex);
            if (targetRow == null)
            {
                targetRow = sheet.CreateRow(rowIndex);
            }

            ICell cell = targetRow.GetCell(colIndex);
            if (cell == null)
            {
                cell = targetRow.CreateCell(colIndex);
            }

            cell.CellStyle = style;

            return cell;
        }

        public static ICell SetCell(ISheet sheet, int rowIndex, int colIndex, string value, ICellStyle style)
        {
            ICell cell = SetCell(sheet, rowIndex, colIndex, value);
            SetCellStyle(sheet, rowIndex, colIndex, style);

            return cell;
        }

        public static void MergeCells(ISheet sheet, int startRowIndex, int endRowIndex, int startColIndex, int endColIndex)
        {
            if (sheet == null || startRowIndex < 0 || endRowIndex < 0 || startColIndex < 0 || endColIndex < 0 || startRowIndex > endRowIndex || startColIndex > endColIndex)
            {
                return;
            }

            CellRangeAddress cellRange = new CellRangeAddress(startRowIndex, endRowIndex, startColIndex, endColIndex);
            sheet.AddMergedRegion(cellRange);
        }

        public static ICellStyle CreateCommonCellStyle(IWorkbook workbook, IFont font, HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment)
        {
            ICellStyle commonStyle = workbook.CreateCellStyle();
            commonStyle.SetFont(font);
            commonStyle.Alignment = horizontalAlignment;
            commonStyle.VerticalAlignment = verticalAlignment;
            commonStyle.BorderTop = BorderStyle.Thin;
            commonStyle.BorderBottom = BorderStyle.Thin;
            commonStyle.BorderLeft = BorderStyle.Thin;
            commonStyle.BorderRight = BorderStyle.Thin;
            commonStyle.WrapText = true;

            return commonStyle;
        }

        public static ICellStyle CreateThinCellStyle(IWorkbook workbook, IFont font, HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment)
        {
            ICellStyle commonStyle = workbook.CreateCellStyle();
            commonStyle.SetFont(font);
            commonStyle.Alignment = horizontalAlignment;
            commonStyle.VerticalAlignment = verticalAlignment;
            commonStyle.BorderTop = BorderStyle.None;
            commonStyle.BorderBottom = BorderStyle.None;
            commonStyle.BorderLeft = BorderStyle.None;
            commonStyle.BorderRight = BorderStyle.None;
            commonStyle.WrapText = true;

            return commonStyle;
        }
    }
}
