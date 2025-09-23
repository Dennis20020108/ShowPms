using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using ShowPms.DTOs;

public class EstimationServices
{
    public byte[] ExportExcel(ExportRequest request)
    {
        XSSFWorkbook workbook = new XSSFWorkbook();
        ISheet sheet = workbook.CreateSheet("估價單");
        int rowIndex = 0;

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
        SetBorder(headerStyle, BorderStyle.Thin, BorderStyle.Double); // 外雙框線

        // ===== 一般儲存格樣式 =====
        ICellStyle cellStyle = workbook.CreateCellStyle();
        cellStyle.SetFont(normalFont);
        SetBorder(cellStyle, BorderStyle.Thin, BorderStyle.Double);

        // ===== 數字儲存格樣式 (靠右) =====
        ICellStyle numberStyle = workbook.CreateCellStyle();
        numberStyle.CloneStyleFrom(cellStyle);
        numberStyle.Alignment = HorizontalAlignment.Right;

        // ===== 標題 =====
        IRow titleRow = sheet.CreateRow(rowIndex++);
        ICell titleCell = titleRow.CreateCell(0);
        titleCell.SetCellValue("秀傳醫療體系工程採購報價單");
        titleCell.CellStyle = titleStyle;
        sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 7));
        titleRow.HeightInPoints = 30;

        // ===== 工程名稱 / 日期 =====
        IRow infoRow = sheet.CreateRow(rowIndex++);
        infoRow.CreateCell(0).SetCellValue("工程名稱：" + request.ProjectName);
        sheet.AddMergedRegion(new CellRangeAddress(1, 1, 0, 1));
        infoRow.CreateCell(6).SetCellValue("日期：    年    月    日");
        sheet.AddMergedRegion(new CellRangeAddress(1, 1, 6, 7));

        // ===== 總表表頭 =====
        IRow header = sheet.CreateRow(rowIndex++);
        string[] headers = { "項次", "工程項目", "數量", "單位", "單價", "金額", "備註", "圖號" };
        for (int i = 0; i < headers.Length; i++)
        {
            ICell h = header.CreateCell(i);
            h.SetCellValue(headers[i]);
            h.CellStyle = headerStyle;
            sheet.SetColumnWidth(i, 15 * 256);
        }

        // ===== 總表 + 明細 (範例) =====
        int sectionNo = 1;
        foreach (var major in request.MajorItems)
        {
            // 大項標題
            IRow majorRow = sheet.CreateRow(rowIndex++);
            majorRow.CreateCell(0).SetCellValue(ToChineseNumber(sectionNo));
            majorRow.GetCell(0).CellStyle = cellStyle;
            majorRow.CreateCell(1).SetCellValue(major.Name);
            majorRow.GetCell(1).CellStyle = cellStyle;

            decimal majorSubTotal = 0;

            foreach (var middle in major.MiddleItems)
            {
                // 中項標題
                IRow middleRow = sheet.CreateRow(rowIndex++);
                middleRow.CreateCell(1).SetCellValue("【" + middle.Name + "】");
                middleRow.GetCell(1).CellStyle = cellStyle;

                decimal middleSubTotal = 0;

                foreach (var item in middle.Items)
                {
                    IRow row = sheet.CreateRow(rowIndex++);

                    row.CreateCell(0).SetCellValue(item.Id); // 或自訂項次
                    row.GetCell(0).CellStyle = cellStyle;

                    row.CreateCell(1).SetCellValue(item.Name);
                    row.GetCell(1).CellStyle = cellStyle;

                    row.CreateCell(2).SetCellValue((double)item.QuantityDecimal);  // ✅ 使用轉換屬性
                    row.GetCell(2).CellStyle = numberStyle;

                    row.CreateCell(3).SetCellValue(item.Unit);
                    row.GetCell(3).CellStyle = cellStyle;

                    row.CreateCell(4).SetCellValue((double)item.UnitPriceDecimal);  // ✅ 使用轉換屬性
                    row.GetCell(4).CellStyle = numberStyle;

                    var amount = item.GetTotalPrice();  // ✅ 使用方法封裝的計算
                    row.CreateCell(5).SetCellValue((double)amount);
                    row.GetCell(5).CellStyle = numberStyle;

                    row.CreateCell(6).SetCellValue(item.Note);
                    row.GetCell(6).CellStyle = cellStyle;

                    row.CreateCell(7).SetCellValue(""); // 若要放圖號，請補欄位
                    row.GetCell(7).CellStyle = cellStyle;

                    middleSubTotal += amount;  // decimal OK
                }


                // 中項小計（可選）
                IRow midSubTotal = sheet.CreateRow(rowIndex++);
                midSubTotal.CreateCell(1).SetCellValue("中項小計");
                midSubTotal.GetCell(1).CellStyle = cellStyle;
                midSubTotal.CreateCell(5).SetCellValue((double)middleSubTotal);
                midSubTotal.GetCell(5).CellStyle = numberStyle;

                majorSubTotal += middleSubTotal;
            }

            // 大項小計
            IRow subTotal = sheet.CreateRow(rowIndex++);
            subTotal.CreateCell(1).SetCellValue("小計");
            subTotal.GetCell(1).CellStyle = cellStyle;
            subTotal.CreateCell(5).SetCellValue((double)majorSubTotal);
            subTotal.GetCell(5).CellStyle = numberStyle;

            sectionNo++;
        }

        using var ms = new MemoryStream();
        workbook.Write(ms);
        return ms.ToArray();
    }

    // 設定內外框線
    private void SetBorder(ICellStyle style, BorderStyle inner, BorderStyle outer)
    {
        style.BorderBottom = inner;
        style.BorderTop = inner;
        style.BorderLeft = inner;
        style.BorderRight = inner;
        // TODO: 外框處理可額外寫判斷 (NPOI 預設無「雙外框」，要用合併樣式或後處理)
    }

    // 阿拉伯數字轉中文大寫 (1 → 一, 2 → 二)
    private string ToChineseNumber(int num)
    {
        string[] map = { "零", "一", "二", "三", "四", "五", "六", "七", "八", "九" };
        return map[num];
    }
}
