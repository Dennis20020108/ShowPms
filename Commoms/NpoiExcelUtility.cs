using NPOI.OOXML.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

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
        /// <summary>
        /// 自適應列高
        /// </summary>
        public static void AutoSizeRowHeight(IRow row, string text, int colWidthInChars, float fontSizePt = 12f)
        {
            if (string.IsNullOrEmpty(text))
            {
                row.HeightInPoints = fontSizePt * 2f;
                return;
            }

            // 假設中文字寬度較英文字大，調整字元寬度估算係數
            int adjustedColWidth = (int)(colWidthInChars * 0.75);

            string[] lines = text.Split(new[] { '\n' }, StringSplitOptions.None);
            int totalLines = 0;

            foreach (var line in lines)
            {
                int lineLength = line.Length;
                int linesNeeded = (int)Math.Ceiling(lineLength / (double)adjustedColWidth);
                totalLines += Math.Max(linesNeeded, 1);
            }

            // 行高乘以字體大小再乘以 2 倍，再加一點額外行距 (5pt)
            row.HeightInPoints = totalLines * fontSizePt * 2f + 5;
        }

        /// <summary>
        /// 所有樣式
        /// </summary>
        public static EstimationStyles CreateEstimationStyles(XSSFWorkbook workbook)
        {
            // ===== 字型定義 =====
            IFont titleFont = workbook.CreateFont();
            titleFont.FontName = "微軟正黑體";
            titleFont.FontHeightInPoints = 20;
            titleFont.IsBold = true;

            IFont headerFont = workbook.CreateFont();
            headerFont.FontName = "微軟正黑體";
            headerFont.FontHeightInPoints = 12;
            headerFont.IsBold = true;

            IFont normalFont = workbook.CreateFont();
            normalFont.FontName = "微軟正黑體";
            normalFont.FontHeightInPoints = 12;

            // ===== 標題樣式 =====
            ICellStyle titleStyle = workbook.CreateCellStyle();
            titleStyle.Alignment = HorizontalAlignment.Center;
            titleStyle.VerticalAlignment = VerticalAlignment.Center;
            titleStyle.SetFont(titleFont);

            // ===== 表頭樣式 =====
            ICellStyle headerStyle = workbook.CreateCellStyle();
            headerStyle.Alignment = HorizontalAlignment.Center;
            headerStyle.VerticalAlignment = VerticalAlignment.Center;
            headerStyle.SetFont(headerFont);
            SetBorder(headerStyle);

            // ===== 一般儲存格樣式（靠左對齊）=====
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.SetFont(normalFont);
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            cellStyle.Alignment = HorizontalAlignment.Left; // 明確設定靠左
            cellStyle.WrapText = true;
            SetBorder(cellStyle);

            // ===== 置中樣式（用於項次、數量） =====
            ICellStyle centerCellStyle = workbook.CreateCellStyle();
            centerCellStyle.CloneStyleFrom(cellStyle);
            centerCellStyle.Alignment = HorizontalAlignment.Center;

            // ===== 數字儲存格樣式 (靠右，整數+千分位) =====
            ICellStyle numberStyle = workbook.CreateCellStyle();
            numberStyle.CloneStyleFrom(cellStyle);
            numberStyle.Alignment = HorizontalAlignment.Right;
            numberStyle.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0");

            // ===== 橘色底色樣式 (#FCD5B4) =====
            XSSFColor orangeColor = new XSSFColor(new byte[] { 252, 213, 180 }, new DefaultIndexedColorMap());

            ICellStyle orangeBgStyle = workbook.CreateCellStyle();
            orangeBgStyle.CloneStyleFrom(cellStyle);
            ((XSSFCellStyle)orangeBgStyle).SetFillForegroundColor(orangeColor);
            orangeBgStyle.FillPattern = FillPattern.SolidForeground;
            orangeBgStyle.Alignment = HorizontalAlignment.Center;

            // ===== 橘色底色 + 數字格式樣式 =====
            ICellStyle orangeBgNumberStyle = workbook.CreateCellStyle();
            orangeBgNumberStyle.CloneStyleFrom(cellStyle);
            ((XSSFCellStyle)orangeBgNumberStyle).SetFillForegroundColor(orangeColor);
            orangeBgNumberStyle.FillPattern = FillPattern.SolidForeground;
            orangeBgNumberStyle.Alignment = HorizontalAlignment.Right;
            orangeBgNumberStyle.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0");

            // ===== 灰色底色樣式 (#D0CECE) =====
            XSSFColor grayColor = new XSSFColor(new byte[] { 208, 206, 206 }, new DefaultIndexedColorMap());

            ICellStyle grayBgStyle = workbook.CreateCellStyle();
            grayBgStyle.CloneStyleFrom(cellStyle);
            ((XSSFCellStyle)grayBgStyle).SetFillForegroundColor(grayColor);
            grayBgStyle.FillPattern = FillPattern.SolidForeground;

            // ===== 灰底 + 置中樣式 =====
            ICellStyle grayCenterStyle = workbook.CreateCellStyle();
            grayCenterStyle.CloneStyleFrom(cellStyle);
            ((XSSFCellStyle)grayCenterStyle).SetFillForegroundColor(grayColor);
            grayCenterStyle.FillPattern = FillPattern.SolidForeground;
            grayCenterStyle.Alignment = HorizontalAlignment.Center;
            grayCenterStyle.VerticalAlignment = VerticalAlignment.Center;

            // ===== 備註專用的無邊框樣式 =====
            ICellStyle noteCellStyle = workbook.CreateCellStyle();
            noteCellStyle.SetFont(normalFont);
            noteCellStyle.VerticalAlignment = VerticalAlignment.Center;
            noteCellStyle.Alignment = HorizontalAlignment.Left;
            noteCellStyle.WrapText = true;

            return new EstimationStyles
            {
                TitleStyle = titleStyle,
                HeaderStyle = headerStyle,
                CellStyle = cellStyle,
                NumberStyle = numberStyle,
                OrangeBgStyle = orangeBgStyle,
                GrayBgStyle = grayBgStyle,
                GrayCenterStyle = grayCenterStyle,
                CenterCellStyle = centerCellStyle,
                OrangeBgNumberStyle = orangeBgNumberStyle,
                NoteCellStyle = noteCellStyle
            };
        }

        /// <summary>
        /// 設定邊框
        /// </summary>
        private static void SetBorder(ICellStyle style)
        {
            style.BorderBottom = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
        }

    }

    /// <summary>
    /// 估價單樣式集合
    /// </summary>
    public class EstimationStyles
    {
        public ICellStyle TitleStyle { get; set; }
        public ICellStyle HeaderStyle { get; set; }
        public ICellStyle CellStyle { get; set; }
        public ICellStyle NumberStyle { get; set; }
        public ICellStyle OrangeBgStyle { get; set; }
        public ICellStyle GrayBgStyle { get; set; }
        public ICellStyle GrayCenterStyle { get; set; }
        public ICellStyle CenterCellStyle { get; set; }
        public ICellStyle OrangeBgNumberStyle { get; set; }
        public ICellStyle NoteCellStyle { get; set; }
    }
}