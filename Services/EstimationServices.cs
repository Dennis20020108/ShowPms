using NPOI.OOXML.XSSF.UserModel;
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

        var styles = CreateStyles(workbook);
        int rowIndex = 0;

        // 設定欄位寬度
        SetColumnWidths(sheet);

        // ===== 標題 =====
        IRow titleRow = sheet.CreateRow(rowIndex++);
        ICell titleCell = titleRow.CreateCell(0);
        titleCell.SetCellValue("秀傳醫療體系工程採購報價單");
        titleCell.CellStyle = styles.titleStyle;
        sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 7));
        titleRow.HeightInPoints = 30;

        // ===== 工程名稱 / 日期 =====
        IRow infoRow = sheet.CreateRow(rowIndex++);
        infoRow.HeightInPoints = 20; //  設定列高為 20pt

        var projectNameCell = infoRow.CreateCell(0);
        projectNameCell.SetCellValue("工程名稱：" + request.ProjectName);
        projectNameCell.CellStyle = styles.cellStyle;
        sheet.AddMergedRegion(new CellRangeAddress(1, 1, 0, 1));

        var dateCell = infoRow.CreateCell(6);
        dateCell.SetCellValue("日期：　　年　　月　　日");
        dateCell.CellStyle = styles.cellStyle;
        sheet.AddMergedRegion(new CellRangeAddress(1, 1, 6, 7));


        // ===== 加入總表區塊 =====
        GenerateSummaryBlock(sheet, request, styles, ref rowIndex);

        // 空一列
        rowIndex++;

        // ===== 加入明細表區塊 =====
        GenerateDetailBlock(sheet, request, styles, ref rowIndex);

        using var ms = new MemoryStream();
        workbook.Write(ms);
        return ms.ToArray();
    }

    private void SetColumnWidths(ISheet sheet)
    {
        // 設定欄位寬度 (單位: 1/256 字符寬度)
        sheet.SetColumnWidth(0, 7.5 * 256);   // A欄：項次 - 較小
        sheet.SetColumnWidth(1, 61.25 * 256);  // B欄：工程項目 - 較寬
        sheet.SetColumnWidth(2, 7.5 * 256);   // C欄：數量 - 較小
        sheet.SetColumnWidth(3, 7.75 * 256);   // D欄：單位 - 較小
        sheet.SetColumnWidth(4, 22 * 256);  // E欄：單價 - 中等
        sheet.SetColumnWidth(5, 23.13 * 256);  // F欄：金額 - 中等
        sheet.SetColumnWidth(6, 28.25 * 256);  // G欄：備註 - 中等
        sheet.SetColumnWidth(7, 14.63 * 256);  // H欄：圖號 - 中等
    }

    private void GenerateSummaryBlock(ISheet sheet, ExportRequest request, (ICellStyle titleStyle, ICellStyle headerStyle, ICellStyle cellStyle, ICellStyle numberStyle, ICellStyle orangeBgStyle, ICellStyle grayBgStyle, ICellStyle grayCenterStyle, ICellStyle centerCellStyle) styles, ref int rowIndex)
    {
        // 總表表頭
        IRow header = sheet.CreateRow(rowIndex++);
        string[] headers = { "項次", "工程項目", "數量", "單位", "單價", "金額", "備註", "圖號" };
        for (int i = 0; i < headers.Length; i++)
        {
            ICell h = header.CreateCell(i);
            h.SetCellValue(headers[i]);
            h.CellStyle = styles.headerStyle;
        }

        int sectionNo = 1;
        foreach (var major in request.MajorItems)
        {
            foreach (var middle in major.MiddleItems)
            {
                IRow row = sheet.CreateRow(rowIndex++);
                row.CreateCell(0).SetCellValue(ToChineseNumber(sectionNo)).CellStyle = styles.centerCellStyle;
                row.CreateCell(1).SetCellValue(middle.Name).CellStyle = styles.cellStyle;
                row.CreateCell(2).SetCellValue(1).CellStyle = styles.centerCellStyle;
                row.CreateCell(3).SetCellValue("式").CellStyle = styles.centerCellStyle;
                row.CreateCell(4).SetCellValue("-").CellStyle = styles.numberStyle;
                row.CreateCell(5).SetCellValue("-").CellStyle = styles.numberStyle;
                row.CreateCell(6).SetCellValue("").CellStyle = styles.cellStyle;
                row.CreateCell(7).SetCellValue("").CellStyle = styles.cellStyle;
                sectionNo++;
            }
        }

        // 小計 - 為每個欄位都加上邊框
        IRow subTotal = sheet.CreateRow(rowIndex++);
        subTotal.CreateCell(0).SetCellValue("").CellStyle = styles.cellStyle;
        subTotal.CreateCell(1).SetCellValue("小計").CellStyle = styles.orangeBgStyle;
        subTotal.CreateCell(2).SetCellValue("").CellStyle = styles.orangeBgStyle;
        subTotal.CreateCell(3).SetCellValue("").CellStyle = styles.orangeBgStyle;
        subTotal.CreateCell(4).SetCellValue("").CellStyle = styles.orangeBgStyle;
        subTotal.CreateCell(5).SetCellValue("-").CellStyle = styles.orangeBgStyle;
        subTotal.CreateCell(6).SetCellValue("").CellStyle = styles.orangeBgStyle;
        subTotal.CreateCell(7).SetCellValue("").CellStyle = styles.cellStyle; // H欄無需上底色

        // 其他費用 
        IRow other = sheet.CreateRow(rowIndex++);
        other.CreateCell(0).SetCellValue(ToChineseNumber(sectionNo)).CellStyle = styles.cellStyle;
        // 應與其他工程項目資料一樣來自資料庫...
        var otherDesc = "1.現場清潔及安全防護,文書資料整理\n" +
                        "2.勞工安全衛生管理費\n" +
                        "3.工程營造綜合保險\n" +
                        "4.工程品管及包商利潤費";

        other.CreateCell(1).SetCellValue(otherDesc);
        other.GetCell(1).CellStyle = styles.cellStyle; // 確保 WrapText = true
        other.CreateCell(2).SetCellValue("6.65").CellStyle = styles.centerCellStyle;
        other.CreateCell(3).SetCellValue("%").CellStyle = styles.centerCellStyle;
        other.CreateCell(4).SetCellValue("").CellStyle = styles.cellStyle;
        other.CreateCell(5).SetCellValue("-").CellStyle = styles.numberStyle;
        other.CreateCell(6).SetCellValue("").CellStyle = styles.cellStyle;
        other.CreateCell(7).SetCellValue("").CellStyle = styles.cellStyle;
        
        // 小計 - 為每個欄位都加上邊框
        IRow subTotal2 = sheet.CreateRow(rowIndex++);
        subTotal2.CreateCell(0).SetCellValue("").CellStyle = styles.cellStyle;
        subTotal2.CreateCell(1).SetCellValue("小計").CellStyle = styles.orangeBgStyle;
        subTotal2.CreateCell(2).SetCellValue("").CellStyle = styles.orangeBgStyle;
        subTotal2.CreateCell(3).SetCellValue("").CellStyle = styles.orangeBgStyle;
        subTotal2.CreateCell(4).SetCellValue("").CellStyle = styles.orangeBgStyle;
        subTotal2.CreateCell(5).SetCellValue("-").CellStyle = styles.orangeBgStyle;
        subTotal2.CreateCell(6).SetCellValue("").CellStyle = styles.orangeBgStyle;
        subTotal2.CreateCell(7).SetCellValue("").CellStyle = styles.cellStyle;

        // 營業稅 - 為每個欄位都加上邊框
        IRow tax = sheet.CreateRow(rowIndex++);
        tax.CreateCell(0).SetCellValue("").CellStyle = styles.cellStyle;
        tax.CreateCell(1).SetCellValue("營業稅").CellStyle = styles.cellStyle;
        tax.CreateCell(2).SetCellValue(5).CellStyle = styles.centerCellStyle;
        tax.CreateCell(3).SetCellValue("%").CellStyle = styles.centerCellStyle;
        tax.CreateCell(4).SetCellValue("").CellStyle = styles.cellStyle;
        tax.CreateCell(5).SetCellValue("-").CellStyle = styles.numberStyle;
        tax.CreateCell(6).SetCellValue("").CellStyle = styles.cellStyle;
        tax.CreateCell(7).SetCellValue("").CellStyle = styles.cellStyle;

        // 總價 - 為每個欄位都加上邊框
        IRow total = sheet.CreateRow(rowIndex++);
        total.CreateCell(0).SetCellValue("").CellStyle = styles.cellStyle;
        total.CreateCell(1).SetCellValue("總價").CellStyle = styles.orangeBgStyle;
        total.CreateCell(2).SetCellValue("").CellStyle = styles.orangeBgStyle;
        total.CreateCell(3).SetCellValue("").CellStyle = styles.orangeBgStyle;
        total.CreateCell(4).SetCellValue("").CellStyle = styles.orangeBgStyle;
        total.CreateCell(5).SetCellValue("").CellStyle = styles.orangeBgStyle;
        total.CreateCell(6).SetCellValue("").CellStyle = styles.orangeBgStyle;
        total.CreateCell(7).SetCellValue("").CellStyle = styles.cellStyle;

        // 附註（多行） - 每行每個欄位都加邊框
        string[] noteHeaders = new string[] { "附註1.", "2.", "3.", "4." };
        string[] noteContents = new string[]
        {
            "付款方式：工程完工支付80%，驗收合格支付17%，保固3%；保固2年(正式驗收合格及工程相關出廠證明資料完備後起算)。",
            "工程保固期間內，如有缺失由承包商負責修繕，含含稅價款計算。",
            "本工程需配合進度加班施工，全部工程合約期間120天。",
            "本工程完工後需附材料防火証明，出廠證明方可辦理驗收。"
        };

        for (int i = 0; i < noteHeaders.Length; i++)
        {
            IRow noteRow = sheet.CreateRow(rowIndex++);
            var headerCell = noteRow.CreateCell(0);
            headerCell.SetCellValue(noteHeaders[i]);
            headerCell.CellStyle = styles.cellStyle;

            var contentCell = noteRow.CreateCell(1);
            contentCell.SetCellValue(noteContents[i]);
            contentCell.CellStyle = styles.cellStyle;
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex - 1, rowIndex - 1, 1, 7));
        }

        // 總價大寫行 - 每個欄位都加邊框
        
        IRow totalInWordsRow = sheet.CreateRow(rowIndex++);

        var totalCell = totalInWordsRow.CreateCell(0);
        totalCell.SetCellValue("總價");
        totalCell.CellStyle = styles.cellStyle;

        var chineseAmountCell = totalInWordsRow.CreateCell(1);
        chineseAmountCell.CellStyle = styles.cellStyle;

        var currencyCell = totalInWordsRow.CreateCell(2);
        currencyCell.SetCellValue("新台幣");
        currencyCell.CellStyle = styles.cellStyle;

        var ntCell = totalInWordsRow.CreateCell(3);
        ntCell.SetCellValue("NT:");
        ntCell.CellStyle = styles.cellStyle;

        var numberAmountCell = totalInWordsRow.CreateCell(4);
        numberAmountCell.CellStyle = styles.numberStyle;
        sheet.AddMergedRegion(new CellRangeAddress(rowIndex - 1, rowIndex - 1, 4, 7));

        // 廠商與用印欄位 - 加邊框
        IRow signRow = sheet.CreateRow(rowIndex++);
        var vendorCell = signRow.CreateCell(0);
        vendorCell.SetCellValue("廠商");
        vendorCell.CellStyle = styles.cellStyle;
        sheet.AddMergedRegion(new CellRangeAddress(rowIndex - 1, rowIndex - 1, 0, 2));

        var stampCell = signRow.CreateCell(3);
        stampCell.SetCellValue("用印");
        stampCell.CellStyle = styles.cellStyle;
        sheet.AddMergedRegion(new CellRangeAddress(rowIndex - 1, rowIndex - 1, 3, 7));

        // 簽章空白區域 - 建立5列空白但有邊框的儲存格
        int signStartRow = rowIndex;
        for (int i = 0; i < 5; i++)
        {
            IRow emptyRow = sheet.CreateRow(rowIndex++);
            for (int j = 0; j < 8; j++)
            {
                var cell = emptyRow.CreateCell(j);
                cell.SetCellValue("");
                cell.CellStyle = styles.cellStyle;
            }
        }

        // 合併簽章區域
        sheet.AddMergedRegion(new CellRangeAddress(signStartRow, signStartRow + 4, 0, 2)); // A-C欄
        sheet.AddMergedRegion(new CellRangeAddress(signStartRow, signStartRow + 4, 3, 7)); // D-H欄
    }

    private string ConvertToChineseCurrency(decimal amount)
    {
        return "零"; // 簡單示範，實際需要完整實作
    }

    private void GenerateDetailBlock(ISheet sheet, ExportRequest request, (ICellStyle titleStyle, ICellStyle headerStyle, ICellStyle cellStyle, ICellStyle numberStyle, ICellStyle orangeBgStyle, ICellStyle grayBgStyle, ICellStyle grayCenterStyle, ICellStyle centerCellStyle) styles, ref int rowIndex)
    {
        // 明細表標頭
        IRow header = sheet.CreateRow(rowIndex++);
        string[] headers = { "項次", "工程項目", "數量", "單位", "單價", "金額", "備註", "圖號" };
        for (int i = 0; i < headers.Length; i++)
        {
            ICell h = header.CreateCell(i);
            h.SetCellValue(headers[i]);
            h.CellStyle = styles.headerStyle;
        }

        int sectionNo = 1;
        foreach (var major in request.MajorItems)
        {
            foreach (var middle in major.MiddleItems)
            {
                // 中項標題列 - 確保所有欄位都有邊框
                IRow middleRow = sheet.CreateRow(rowIndex++);
                middleRow.CreateCell(0).SetCellValue(ToChineseNumber(sectionNo)).CellStyle = styles.grayCenterStyle;
                middleRow.CreateCell(1).SetCellValue(middle.Name).CellStyle = styles.grayBgStyle;

                // 為其他欄位也加上空白內容和邊框
                for (int i = 2; i < 8; i++)
                {
                    middleRow.CreateCell(i).SetCellValue("").CellStyle = styles.cellStyle;
                }

                int itemIndex = 1;
                foreach (var item in middle.Items)
                {
                    IRow row = sheet.CreateRow(rowIndex++);
                    row.CreateCell(0).SetCellValue(itemIndex++).CellStyle = styles.centerCellStyle;
                    row.CreateCell(1).SetCellValue(item.Name).CellStyle = styles.cellStyle;
                    row.CreateCell(2).SetCellValue((double)item.QuantityDecimal).CellStyle = styles.centerCellStyle;
                    row.CreateCell(3).SetCellValue(item.Unit).CellStyle = styles.centerCellStyle;
                    row.CreateCell(4).SetCellValue("").CellStyle = styles.numberStyle;
                    row.CreateCell(5).SetCellValue("").CellStyle = styles.numberStyle;
                    row.CreateCell(6).SetCellValue(item.Note).CellStyle = styles.cellStyle;
                    row.CreateCell(7).SetCellValue("").CellStyle = styles.cellStyle;
                }

                // 小計 - 確保所有欄位都有邊框
                IRow subtotalRow = sheet.CreateRow(rowIndex++);
                subtotalRow.CreateCell(0).SetCellValue("").CellStyle = styles.orangeBgStyle;
                subtotalRow.CreateCell(1).SetCellValue("小計").CellStyle = styles.orangeBgStyle;
                subtotalRow.CreateCell(2).SetCellValue("").CellStyle = styles.orangeBgStyle;
                subtotalRow.CreateCell(3).SetCellValue("").CellStyle = styles.orangeBgStyle;
                subtotalRow.CreateCell(4).SetCellValue("").CellStyle = styles.orangeBgStyle;
                subtotalRow.CreateCell(5).SetCellValue("").CellStyle = styles.orangeBgStyle;
                subtotalRow.CreateCell(6).SetCellValue("").CellStyle = styles.orangeBgStyle;
                subtotalRow.CreateCell(7).SetCellValue("").CellStyle = styles.orangeBgStyle;

                // 空白列（有邊框）
                IRow blankRow = sheet.CreateRow(rowIndex++);
                for (int i = 0; i < 8; i++)
                {
                    var cell = blankRow.CreateCell(i);
                    cell.SetCellValue("");
                    cell.CellStyle = styles.cellStyle;
                }

                sectionNo++;
            }
        }
    }

    private (ICellStyle titleStyle, ICellStyle headerStyle, ICellStyle cellStyle, ICellStyle numberStyle, ICellStyle orangeBgStyle, ICellStyle grayBgStyle, ICellStyle grayCenterStyle, ICellStyle centerCellStyle) CreateStyles(XSSFWorkbook workbook)
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
        // 標題不加邊框

        // ===== 表頭樣式 =====
        ICellStyle headerStyle = workbook.CreateCellStyle();
        headerStyle.Alignment = HorizontalAlignment.Center;
        headerStyle.VerticalAlignment = VerticalAlignment.Center;
        headerStyle.SetFont(headerFont);
        SetBorder(headerStyle);

        // ===== 一般儲存格樣式 =====
        ICellStyle cellStyle = workbook.CreateCellStyle();
        cellStyle.SetFont(normalFont);
        cellStyle.VerticalAlignment = VerticalAlignment.Center;
        SetBorder(cellStyle);
        cellStyle.WrapText = true;

        // ===== 置中樣式（用於項次、數量） =====
        ICellStyle centerCellStyle = workbook.CreateCellStyle();
        centerCellStyle.CloneStyleFrom(cellStyle);
        centerCellStyle.Alignment = HorizontalAlignment.Center;

        // ===== 數字儲存格樣式 (靠右) =====
        ICellStyle numberStyle = workbook.CreateCellStyle();
        numberStyle.CloneStyleFrom(cellStyle);
        numberStyle.Alignment = HorizontalAlignment.Right;

        // ===== 橘色底色樣式 (#FCD5B4) =====
        ICellStyle orangeBgStyle = workbook.CreateCellStyle();
        orangeBgStyle.CloneStyleFrom(cellStyle);
        XSSFColor orangeColor = new XSSFColor(new byte[] { 252, 213, 180 }, new DefaultIndexedColorMap());
        ((XSSFCellStyle)orangeBgStyle).SetFillForegroundColor(orangeColor);
        orangeBgStyle.FillPattern = FillPattern.SolidForeground;
        orangeBgStyle.Alignment = HorizontalAlignment.Center;

        // 先定義 grayColor
        XSSFColor grayColor = new XSSFColor(new byte[] { 208, 206, 206 }, new DefaultIndexedColorMap());

        // ===== 灰色底色樣式 (#D0CECE) =====
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


        return (titleStyle, headerStyle, cellStyle, numberStyle, orangeBgStyle, grayBgStyle, grayCenterStyle, centerCellStyle);
    }

    // 設定邊框
    private void SetBorder(ICellStyle style)
    {
        style.BorderBottom = BorderStyle.Thin;
        style.BorderTop = BorderStyle.Thin;
        style.BorderLeft = BorderStyle.Thin;
        style.BorderRight = BorderStyle.Thin;
    }

    // 阿拉伯數字轉中文大寫 (1 → 一, 2 → 二)
    private string ToChineseNumber(int num)
    {
        string[] digits = { "零", "一", "二", "三", "四", "五", "六", "七", "八", "九" };

        if (num < 0)
            return num.ToString(); // 防呆處理：負數直接返回阿拉伯數字

        if (num < 10)
            return digits[num];

        if (num == 10)
            return "十";

        if (num < 20)
            return "十" + digits[num % 10];

        if (num < 100)
        {
            int ten = num / 10;
            int one = num % 10;

            string result = digits[ten] + "十";

            if (one != 0)
                result += digits[one];

            return result;
        }

        return num.ToString(); // 超過99就直接顯示阿拉伯數字
    }


}