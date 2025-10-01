using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using ShowPms.Commoms;
using ShowPms.DTOs;

public class EstimationServices
{
    private Dictionary<int, int> detailSubtotalRowMap = new Dictionary<int, int>();

    public byte[] ExportExcel(ExportRequest request)
    {
        XSSFWorkbook workbook = new XSSFWorkbook();
        ISheet sheet = workbook.CreateSheet("估價單");
        ISheet sheet2 = workbook.CreateSheet("結構體及外牆裝修 備註");
        ISheet sheet3 = workbook.CreateSheet("裝修 備註");

        var styles = NpoiExcelUtility.CreateEstimationStyles(workbook);
        int rowIndex = 0;

        SetColumnWidths(sheet);

        // ===== 標題 =====
        IRow titleRow = sheet.CreateRow(rowIndex++);
        ICell titleCell = titleRow.CreateCell(0);
        titleCell.SetCellValue("秀傳醫療體系工程採購報價單");
        titleCell.CellStyle = styles.TitleStyle;
        sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 7));
        titleRow.HeightInPoints = 30;

        // ===== 工程名稱 / 日期 =====
        IRow infoRow = sheet.CreateRow(rowIndex++);
        infoRow.HeightInPoints = 20;

        var projectNameCell = infoRow.CreateCell(0);
        projectNameCell.SetCellValue("工程名稱：" + request.ProjectName);
        projectNameCell.CellStyle = styles.CellStyle;
        sheet.AddMergedRegion(new CellRangeAddress(1, 1, 0, 1));

        var dateCell = infoRow.CreateCell(6);
        dateCell.SetCellValue("製表日期：　　年　　月　　日");
        dateCell.CellStyle = styles.CellStyle;
        sheet.AddMergedRegion(new CellRangeAddress(1, 1, 6, 7));

        // 記錄總表開始的位置
        int summaryHeaderRowIndex = rowIndex;

        // ===== 加入總表區塊（第一次先不加公式）=====
        GenerateSummaryBlock(sheet, request, styles, ref rowIndex, null);

        // 空一列
        rowIndex++;

        // ===== 加入明細表區塊，並收集小計行號 =====
        detailSubtotalRowMap.Clear();
        GenerateDetailBlock(sheet, request, styles, ref rowIndex);

        // ===== 回填總表的複價公式 =====
        FillSummaryPriceFormulas(sheet, request, styles);

        // ===== 計算並填入中文大寫金額 =====
        try
        {
            // 先評估所有公式
            var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
            evaluator.EvaluateAll();

            // 計算總表各行位置
            int totalItems = 0;
            foreach (var major in request.MajorItems)
            {
                totalItems += major.MiddleItems.Count;
            }

            int firstDataRow = summaryHeaderRowIndex + 2;
            int finalTotalRow = firstDataRow + totalItems + 5;
            int totalInWordsRowIndex = finalTotalRow + 22;

            // 取得總價值
            IRow finalTotalRowObj = sheet.GetRow(finalTotalRow - 1);
            ICell finalTotalCell = finalTotalRowObj?.GetCell(5);

            if (finalTotalCell != null)
            {
                var cellValue = evaluator.Evaluate(finalTotalCell);
                double totalAmount = cellValue.NumberValue;

                IRow totalInWordsRow = sheet.GetRow(totalInWordsRowIndex - 1);
                ICell chineseAmountCell = totalInWordsRow?.GetCell(1);

                if (chineseAmountCell != null)
                {
                    chineseAmountCell.SetCellValue(ConvertToChineseCurrency((decimal)totalAmount));
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"計算中文金額失敗: {ex.Message}");
        }

        // ===== 新增：產生備註頁面 =====
        GenerateNoteSheet(sheet2, "結構體及外牆裝修", styles, request);
        GenerateNoteSheet(sheet3, "裝修", styles, request);

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

    // 空白估價單總表
    private void GenerateSummaryBlock(ISheet sheet, ExportRequest request,
        EstimationStyles styles, ref int rowIndex, Dictionary<int, int> detailSubtotalRowMap)
    {
        // 總表表頭
        IRow header = sheet.CreateRow(rowIndex++);
        string[] headers = { "項次", "工程項目", "數量", "單位", "單價", "複價", "備註", "圖號" };
        for (int i = 0; i < headers.Length; i++)
        {
            ICell h = header.CreateCell(i);
            h.SetCellValue(headers[i]);
            h.CellStyle = styles.HeaderStyle;
        }

        int sectionNo = 1;
        int totalItems = 0;
        int firstRowIndex = rowIndex;

        foreach (var major in request.MajorItems)
        {
            foreach (var middle in major.MiddleItems)
            {
                IRow row = sheet.CreateRow(rowIndex++);
                row.CreateCell(0).SetCellValue(ToChineseNumber(sectionNo)).CellStyle = styles.CenterCellStyle;
                row.CreateCell(1).SetCellValue(middle.Name).CellStyle = styles.CellStyle;
                row.CreateCell(2).SetCellValue(1).CellStyle = styles.CenterCellStyle;
                row.CreateCell(3).SetCellValue("式").CellStyle = styles.CenterCellStyle;
                row.CreateCell(4).SetCellValue("-").CellStyle = styles.NumberStyle;

                // 複價欄位先留空，之後會填入公式
                ICell priceCell = row.CreateCell(5);
                priceCell.SetCellValue("-");
                priceCell.CellStyle = styles.NumberStyle;

                row.CreateCell(6).SetCellValue("").CellStyle = styles.CellStyle;
                row.CreateCell(7).SetCellValue("").CellStyle = styles.CellStyle;
                sectionNo++;
                totalItems++;
            }
        }

        // 動態設置小計的文字，將項數轉換為中文數字
        string subTotalText = $"小計（一～{ToChineseNumber(totalItems)}項）";

        // 小計 - 為每個欄位都加上邊框
        IRow subTotal = sheet.CreateRow(rowIndex++);
        subTotal.CreateCell(0).SetCellValue("").CellStyle = styles.CellStyle;
        subTotal.CreateCell(1).SetCellValue(subTotalText).CellStyle = styles.OrangeBgStyle;
        subTotal.CreateCell(2).SetCellValue("").CellStyle = styles.OrangeBgStyle;
        subTotal.CreateCell(3).SetCellValue("").CellStyle = styles.OrangeBgStyle;
        subTotal.CreateCell(4).SetCellValue("").CellStyle = styles.OrangeBgStyle;
        subTotal.CreateCell(5).SetCellValue("-").CellStyle = styles.OrangeBgNumberStyle;
        subTotal.CreateCell(6).SetCellValue("").CellStyle = styles.OrangeBgStyle;
        subTotal.CreateCell(7).SetCellValue("").CellStyle = styles.CellStyle;

        // 其他費用
        IRow other = sheet.CreateRow(rowIndex++);
        other.CreateCell(0).SetCellValue(ToChineseNumber(sectionNo)).CellStyle = styles.CenterCellStyle;
        var otherDesc = "1.現場清潔及安全防護,文書資料整理\n" +
                        "2.勞工安全衛生管理費\n" +
                        "3.工程營造綜合保險\n" +
                        "4.工程品管及包商利潤費";

        other.CreateCell(1).SetCellValue(otherDesc);
        other.GetCell(1).CellStyle = styles.CellStyle;
        other.CreateCell(2).SetCellValue("6.65").CellStyle = styles.CenterCellStyle;
        other.CreateCell(3).SetCellValue("%").CellStyle = styles.CenterCellStyle;
        other.CreateCell(4).SetCellValue("-").CellStyle = styles.CellStyle;
        other.CreateCell(5).SetCellValue("-").CellStyle = styles.NumberStyle;
        other.CreateCell(6).SetCellValue("").CellStyle = styles.CellStyle;
        other.CreateCell(7).SetCellValue("").CellStyle = styles.CellStyle;

        totalItems++;

        // 動態設置合計的文字，將項數轉換為中文數字
        string totalText = $"合計（一～{ToChineseNumber(totalItems)}項）";

        // 合計 - 為每個欄位都加上邊框
        IRow subTotal2 = sheet.CreateRow(rowIndex++);
        subTotal2.CreateCell(0).SetCellValue("").CellStyle = styles.CellStyle;
        subTotal2.CreateCell(1).SetCellValue(totalText).CellStyle = styles.OrangeBgStyle;
        subTotal2.CreateCell(2).SetCellValue("").CellStyle = styles.OrangeBgStyle;
        subTotal2.CreateCell(3).SetCellValue("").CellStyle = styles.OrangeBgStyle;
        subTotal2.CreateCell(4).SetCellValue("").CellStyle = styles.OrangeBgStyle;
        subTotal2.CreateCell(5).SetCellValue("-").CellStyle = styles.OrangeBgNumberStyle;
        subTotal2.CreateCell(6).SetCellValue("").CellStyle = styles.OrangeBgStyle;
        subTotal2.CreateCell(7).SetCellValue("").CellStyle = styles.CellStyle;

        // 營業稅 - 為每個欄位都加上邊框
        IRow tax = sheet.CreateRow(rowIndex++);
        tax.CreateCell(0).SetCellValue("").CellStyle = styles.CellStyle;
        tax.CreateCell(1).SetCellValue("營業稅").CellStyle = styles.CellStyle;
        tax.CreateCell(2).SetCellValue(5).CellStyle = styles.CenterCellStyle;
        tax.CreateCell(3).SetCellValue("%").CellStyle = styles.CenterCellStyle;
        tax.CreateCell(4).SetCellValue("-").CellStyle = styles.CellStyle;
        tax.CreateCell(5).SetCellValue("-").CellStyle = styles.NumberStyle;
        tax.CreateCell(6).SetCellValue("").CellStyle = styles.CellStyle;
        tax.CreateCell(7).SetCellValue("").CellStyle = styles.CellStyle;

        // 總價 - 為每個欄位都加上邊框
        IRow total = sheet.CreateRow(rowIndex++);
        total.CreateCell(0).SetCellValue("").CellStyle = styles.CellStyle;
        total.CreateCell(1).SetCellValue("總價").CellStyle = styles.OrangeBgStyle;
        total.CreateCell(2).SetCellValue("").CellStyle = styles.OrangeBgStyle;
        total.CreateCell(3).SetCellValue("").CellStyle = styles.OrangeBgStyle;
        total.CreateCell(4).SetCellValue("").CellStyle = styles.OrangeBgStyle;
        total.CreateCell(5).SetCellValue("").CellStyle = styles.OrangeBgNumberStyle;
        total.CreateCell(6).SetCellValue("").CellStyle = styles.OrangeBgStyle;
        total.CreateCell(7).SetCellValue("").CellStyle = styles.CellStyle;

        // 插入公式部分
        int summaryHeaderRow = firstRowIndex - 1;
        int firstDataRow = firstRowIndex + 1;
        int lastDataRow = firstDataRow + totalItems - 2;

        int subTotalRow = firstRowIndex + totalItems;
        int otherFeeRow = subTotalRow + 1;
        int totalRow = otherFeeRow + 1;
        int taxRow = totalRow + 1;
        int finalTotalRow = taxRow + 1;

        // 小計複價公式：SUM(F4:F20)
        ICell subTotalValueCell = subTotal.GetCell(5) ?? subTotal.CreateCell(5);
        subTotalValueCell.SetCellFormula($"SUM(F{firstDataRow}:F{lastDataRow})");
        subTotalValueCell.CellStyle = styles.OrangeBgNumberStyle;

        // 其他費用複價公式：=F21*C22/100
        ICell otherValueCell = other.GetCell(5) ?? other.CreateCell(5);
        otherValueCell.SetCellFormula($"F{subTotalRow}*C{otherFeeRow}/100");
        otherValueCell.CellStyle = styles.NumberStyle;

        // 合計複價公式：=F21+F22
        ICell totalValueCell = subTotal2.GetCell(5) ?? subTotal2.CreateCell(5);
        totalValueCell.SetCellFormula($"F{subTotalRow}+F{otherFeeRow}");
        totalValueCell.CellStyle = styles.OrangeBgNumberStyle;

        // 營業稅複價公式：=F23*C24/100
        ICell taxValueCell = tax.GetCell(5) ?? tax.CreateCell(5);
        taxValueCell.SetCellFormula($"F{totalRow}*C{taxRow}/100");
        taxValueCell.CellStyle = styles.NumberStyle;

        // 總價複價公式：=SUM(F23:F24) 或 =F23+F24
        ICell finalTotalValueCell = total.GetCell(5) ?? total.CreateCell(5);
        finalTotalValueCell.SetCellFormula($"F{totalRow}+F{taxRow}");
        finalTotalValueCell.CellStyle = styles.OrangeBgNumberStyle;

        // 附註（多行） - 每行每個欄位都加邊框
        string[] noteHeaders = new string[] { "附註1.", "2.", "3.", "4.", "5.", "6.", "7.", "8.", "9.", "10.", "11.", "12.", "13.", "14.", "15.", "16.", "17.", "18.", "19.", "20.", "21." };
        string[] noteContents = new string[]
        {
            "付款方式：工程完工支付80%,驗收合格支付17%，保固3%，保固2年(正式驗收合格及工程相關出廠證明資料完備後起算保固)，工程依合約完工，功能測試正常，現場驗收合格，提供材料出廠或進口完稅證明，\r\n                  合約約定之書面證明資料及圖面，並檢附保固書稿，工程結算書稿，施工前及施工後現場相片，供本院內部 \"工程驗收報告單\"簽核，每期請款:每月20日前送估驗單，次月月底付款。",
            "工程拆下之廢料由承包商負責清運，至合法棄置場所，可用之器具負責運至指定地點存放。",
            "本工程需配合進度加班趕工，全部工期含備料     120    天。",
            "本工程完工後需附材料防火証明，出廠証明方可辦理驗收，施工前須提供施工進度表。",
            "本工地垃圾需每日下班前清潔，未清潔部份秀傳代為清潔，每一人次扣工程款新台幣貳仟元整。",
            "本工程施工車輛除卸貨外須停放於劃線停車格，任意停車經查屬實每車每日扣款新台幣參仟元整。",
            "本工程施工人員需換發施工識別證，如未穿戴每一人次扣工程款新台幣壹仟元整。",
            "本工程廠商需遵守環保及勞工安全衛生法令規定，如未依規定受罰均由乙方負責。",
            "本工程廠商需勘查現場，如報價有誤，由乙方自行吸收(未勘查現場者，以總價承攬)。",
            "本工程廠商須自備垃圾桶，垃圾須隨時清潔，不得任意放置,違者拍照存證每一廢棄物扣款新台幣伍百元整。",
            "本工程廠商需評估耗材及損料，驗收時則依現場實作實算，追加減處理。",
            "甲方支付工程營造綜合保險費用，未加保並開立證明者工程款中扣除，如工程發生任何意外均由乙方負責。",
            "本工程合約書由乙方製作正本貳份，雙方印花稅由乙方貼足，副本二份(30萬(含)以下免合約)。",
            "本工程採實做實算。",
            "須配合室內裝修審查之施工及作業要求。",
            "如無遇不可抗拒之因素而延宕工期，每日罰款10,000元(有合約者依合約規定辦理)。",
            "本工程為假日施工。",
            "由工務配置獨立電源提供施工期間使用。",
            "須配合室內裝修審查出據相關證明。",
            "工程現場配合人員：楊歲興，電話：0975-619122，FAX：06-2607672",
            "採購處承辦者：李蓮川，電話： 0975-619763，FAX：06-2884832",
        };

        for (int i = 0; i < noteHeaders.Length; i++)
        {
            IRow noteRow = sheet.CreateRow(rowIndex++);
            var headerCell = noteRow.CreateCell(0);
            headerCell.SetCellValue(noteHeaders[i]);
            headerCell.CellStyle = styles.CellStyle;

            var contentCell = noteRow.CreateCell(1);
            contentCell.SetCellValue(noteContents[i]);
            contentCell.CellStyle = styles.CellStyle;
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex - 1, rowIndex - 1, 1, 7));
        }

        // 總價大寫行 - 每個欄位都加邊框
        IRow totalInWordsRow = sheet.CreateRow(rowIndex++);
        var totalCell = totalInWordsRow.CreateCell(0);
        totalCell.SetCellValue("總價");
        totalCell.CellStyle = styles.CellStyle;

        var chineseAmountCell = totalInWordsRow.CreateCell(1);
        // 中文金額先留空，稍後在 ExportExcel 方法最後統一計算
        chineseAmountCell.SetCellValue("");
        chineseAmountCell.CellStyle = styles.CellStyle;

        var currencyCell = totalInWordsRow.CreateCell(2);
        currencyCell.SetCellValue("新台幣");
        currencyCell.CellStyle = styles.CellStyle;

        var ntCell = totalInWordsRow.CreateCell(3);
        ntCell.SetCellValue("NT$:");
        ntCell.CellStyle = styles.CellStyle;

        var numberAmountCell = totalInWordsRow.CreateCell(4);
        numberAmountCell.SetCellFormula($"TEXT(F{finalTotalRow},\"#,##0\")");
        numberAmountCell.CellStyle = styles.NumberStyle;
        sheet.AddMergedRegion(new CellRangeAddress(rowIndex - 1, rowIndex - 1, 4, 7));

        // 廠商與用印欄位 - 加邊框
        IRow signRow = sheet.CreateRow(rowIndex++);
        var vendorCell = signRow.CreateCell(0);
        vendorCell.SetCellValue("廠商");
        vendorCell.CellStyle = styles.CellStyle;
        sheet.AddMergedRegion(new CellRangeAddress(rowIndex - 1, rowIndex - 1, 0, 2));

        var stampCell = signRow.CreateCell(3);
        stampCell.SetCellValue("用印");
        stampCell.CellStyle = styles.CellStyle;
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
                cell.CellStyle = styles.CellStyle;
            }
        }

        // 合併簽章區域
        sheet.AddMergedRegion(new CellRangeAddress(signStartRow, signStartRow + 4, 0, 2)); // A-C欄
        sheet.AddMergedRegion(new CellRangeAddress(signStartRow, signStartRow + 4, 3, 7)); // D-H欄
    }

    // 空白估價單明細表
    private void GenerateDetailBlock(ISheet sheet, ExportRequest request,
        EstimationStyles styles, ref int rowIndex)
    {
        // 明細表標頭
        IRow header = sheet.CreateRow(rowIndex++);
        string[] headers = { "項次", "工程項目", "數量", "單位", "單價", "複價", "備註", "圖號" };
        for (int i = 0; i < headers.Length; i++)
        {
            ICell h = header.CreateCell(i);
            h.SetCellValue(headers[i]);
            h.CellStyle = styles.HeaderStyle;
        }

        int sectionNo = 1;
        int middleItemIndex = 0; // 用來對應總表中的中項索引

        foreach (var major in request.MajorItems)
        {
            foreach (var middle in major.MiddleItems)
            {
                // 中項標題列 - 確保所有欄位都有邊框
                IRow middleRow = sheet.CreateRow(rowIndex++);
                middleRow.CreateCell(0).SetCellValue(ToChineseNumber(sectionNo)).CellStyle = styles.GrayCenterStyle;
                middleRow.CreateCell(1).SetCellValue(middle.Name).CellStyle = styles.GrayBgStyle;

                // 為其他欄位也加上空白內容和邊框
                for (int i = 2; i < 8; i++)
                {
                    middleRow.CreateCell(i).SetCellValue("").CellStyle = styles.CellStyle;
                }

                // 記錄明細項目的起始行（用於小計公式）
                int detailStartRow = rowIndex + 1; // Excel 行號（從 1 開始）

                int itemIndex = 1;
                foreach (var item in middle.Items)
                {
                    IRow row = sheet.CreateRow(rowIndex++);
                    int currentExcelRow = rowIndex; // 當前 Excel 行號

                    row.CreateCell(0).SetCellValue(itemIndex++).CellStyle = styles.CenterCellStyle;
                    row.CreateCell(1).SetCellValue(item.Name).CellStyle = styles.CellStyle;
                    row.CreateCell(2).SetCellValue((double)item.QuantityDecimal).CellStyle = styles.CenterCellStyle;
                    row.CreateCell(3).SetCellValue(item.Unit).CellStyle = styles.CenterCellStyle;
                    row.CreateCell(4).SetCellValue("").CellStyle = styles.NumberStyle;

                    // 複價公式：=C*E (數量*單價)
                    ICell priceCell = row.CreateCell(5);
                    priceCell.SetCellFormula($"C{currentExcelRow}*E{currentExcelRow}");
                    priceCell.CellStyle = styles.NumberStyle;

                    row.CreateCell(6).SetCellValue(item.Note).CellStyle = styles.CellStyle;
                    row.CreateCell(7).SetCellValue("").CellStyle = styles.CellStyle;
                }

                int detailEndRow = rowIndex; // 最後一個明細項目的 Excel 行號

                // 小計 - 確保所有欄位都有邊框
                IRow subtotalRow = sheet.CreateRow(rowIndex++);
                int subtotalExcelRow = rowIndex; // 小計的 Excel 行號

                // 記錄這個中項的小計行號到 map 中
                detailSubtotalRowMap[middleItemIndex] = subtotalExcelRow;
                middleItemIndex++;

                subtotalRow.CreateCell(0).SetCellValue("").CellStyle = styles.OrangeBgStyle;
                subtotalRow.CreateCell(1).SetCellValue("小計").CellStyle = styles.OrangeBgStyle;
                subtotalRow.CreateCell(2).SetCellValue("").CellStyle = styles.OrangeBgStyle;
                subtotalRow.CreateCell(3).SetCellValue("").CellStyle = styles.OrangeBgStyle;
                subtotalRow.CreateCell(4).SetCellValue("").CellStyle = styles.OrangeBgStyle;

                // 小計公式：=SUM(F40:F46)
                ICell subtotalPriceCell = subtotalRow.CreateCell(5);
                subtotalPriceCell.SetCellFormula($"SUM(F{detailStartRow}:F{detailEndRow})");
                subtotalPriceCell.CellStyle = styles.OrangeBgNumberStyle;

                subtotalRow.CreateCell(6).SetCellValue("").CellStyle = styles.OrangeBgStyle;
                subtotalRow.CreateCell(7).SetCellValue("").CellStyle = styles.OrangeBgStyle;

                // 空白列（有邊框）
                IRow blankRow = sheet.CreateRow(rowIndex++);
                for (int i = 0; i < 8; i++)
                {
                    var cell = blankRow.CreateCell(i);
                    cell.SetCellValue("");
                    cell.CellStyle = styles.CellStyle;
                }

                sectionNo++;
            }
        }
    }

    // 計算公式
    private void FillSummaryPriceFormulas(ISheet sheet, ExportRequest request,
        EstimationStyles styles)
    {
        // 找到總表的起始位置（表頭後的第一行數據）
        int summaryDataStartRow = 4; // Excel 行號（假設從第4行開始）

        int itemIndex = 0;
        foreach (var major in request.MajorItems)
        {
            foreach (var middle in major.MiddleItems)
            {
                int currentSummaryRow = summaryDataStartRow + itemIndex;

                // 從 map 中取得對應的明細小計行號
                if (detailSubtotalRowMap.ContainsKey(itemIndex))
                {
                    int detailSubtotalRow = detailSubtotalRowMap[itemIndex];

                    IRow row = sheet.GetRow(currentSummaryRow - 1); // -1 因為 rowIndex 從 0 開始
                    ICell priceCell = row.GetCell(5) ?? row.CreateCell(5);
                    priceCell.SetCellFormula($"F{detailSubtotalRow}");
                    priceCell.CellStyle = styles.NumberStyle;
                }

                itemIndex++;
            }
        }

        // 計算總表的其他公式
        int totalItems = itemIndex;
        int firstDataRow = summaryDataStartRow;
        int lastDataRow = firstDataRow + totalItems - 1;

        int subTotalRow = lastDataRow + 1;
        int otherFeeRow = subTotalRow + 1;
        int totalRow = otherFeeRow + 1;
        int taxRow = totalRow + 1;
        int finalTotalRow = taxRow + 1;

        // 小計公式
        IRow subTotalRowObj = sheet.GetRow(subTotalRow - 1);
        ICell subTotalValueCell = subTotalRowObj.GetCell(5) ?? subTotalRowObj.CreateCell(5);
        subTotalValueCell.SetCellFormula($"SUM(F{firstDataRow}:F{lastDataRow})");
        subTotalValueCell.CellStyle = styles.OrangeBgNumberStyle;

        // 其他費用公式
        IRow otherRowObj = sheet.GetRow(otherFeeRow - 1);
        ICell otherValueCell = otherRowObj.GetCell(5) ?? otherRowObj.CreateCell(5);
        otherValueCell.SetCellFormula($"F{subTotalRow}*C{otherFeeRow}/100");
        otherValueCell.CellStyle = styles.NumberStyle;

        // 合計公式
        IRow totalRowObj = sheet.GetRow(totalRow - 1);
        ICell totalValueCell = totalRowObj.GetCell(5) ?? totalRowObj.CreateCell(5);
        totalValueCell.SetCellFormula($"F{subTotalRow}+F{otherFeeRow}");
        totalValueCell.CellStyle = styles.OrangeBgNumberStyle;

        // 營業稅公式
        IRow taxRowObj = sheet.GetRow(taxRow - 1);
        ICell taxValueCell = taxRowObj.GetCell(5) ?? taxRowObj.CreateCell(5);
        taxValueCell.SetCellFormula($"F{totalRow}*C{taxRow}/100");
        taxValueCell.CellStyle = styles.NumberStyle;

        // 總價公式
        IRow finalTotalRowObj = sheet.GetRow(finalTotalRow - 1);
        ICell finalTotalValueCell = finalTotalRowObj.GetCell(5) ?? finalTotalRowObj.CreateCell(5);
        finalTotalValueCell.SetCellFormula($"F{totalRow}+F{taxRow}");
        finalTotalValueCell.CellStyle = styles.OrangeBgNumberStyle;
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

    /// <summary>
    /// 將數字金額轉換為中文大寫金額
    /// </summary>
    /// <param name="amount">金額數字</param>
    /// <returns>中文大寫金額字串</returns>
    private string ConvertToChineseCurrency(decimal amount)
    {
        if (amount == 0)
            return "零元整";

        // 處理負數
        string prefix = "";
        if (amount < 0)
        {
            prefix = "負";
            amount = Math.Abs(amount);
        }

        // 中文數字
        string[] numbers = { "零", "壹", "貳", "參", "肆", "伍", "陸", "柒", "捌", "玖" };
        // 單位
        string[] units = { "", "拾", "佰", "仟" };
        string[] bigUnits = { "", "萬", "億", "兆" };

        // 分離整數部分
        long integerPart = (long)Math.Floor(amount);

        string result = "";

        // 處理整數部分
        if (integerPart == 0)
        {
            result = "零元";
        }
        else
        {
            string integerStr = integerPart.ToString();
            int length = integerStr.Length;

            // 從高位到低位處理
            for (int i = 0; i < length; i++)
            {
                int digit = int.Parse(integerStr[i].ToString());
                int position = length - i - 1; // 當前位置（從右往左數）
                int unitIndex = position % 4; // 個十百千的位置
                int bigUnitIndex = position / 4; // 萬億兆的位置

                if (digit == 0)
                {
                    // 處理零的特殊情況
                    // 如果不是最後一位，且後面還有非零數字，才加零
                    if (position > 0 && !result.EndsWith("零"))
                    {
                        // 檢查後面是否還有非零數字
                        bool hasNonZeroAfter = false;
                        for (int j = i + 1; j < length && (length - j - 1) / 4 == bigUnitIndex; j++)
                        {
                            if (int.Parse(integerStr[j].ToString()) != 0)
                            {
                                hasNonZeroAfter = true;
                                break;
                            }
                        }
                        if (hasNonZeroAfter)
                        {
                            result += numbers[0];
                        }
                    }
                }
                else
                {
                    result += numbers[digit] + units[unitIndex];
                }

                // 添加大單位（萬、億、兆）
                if (unitIndex == 0 && bigUnitIndex > 0)
                {
                    // 檢查這一節是否全為零
                    bool isAllZero = true;
                    for (int j = Math.Max(0, i - 3); j <= i; j++)
                    {
                        if (int.Parse(integerStr[j].ToString()) != 0)
                        {
                            isAllZero = false;
                            break;
                        }
                    }

                    if (!isAllZero)
                    {
                        result += bigUnits[bigUnitIndex];
                    }
                }
            }

            result += "元";
        }

        // 只需要整數部分，不處理角和分
        return prefix + result;
    }

    // 產生備註頁面
    private void GenerateNoteSheet(ISheet sheet, string sheetType,
        EstimationStyles styles, ExportRequest request)
    {
        int rowIndex = 2; // Excel 的 A3 是 index = 2（從0開始）

        // 設定欄位寬度
        SetNoteSheetColumnWidths(sheet);

        // ===== 標題：工程估價單（A3~G3）=====
        IRow titleRow = sheet.CreateRow(rowIndex++);
        titleRow.HeightInPoints = 30;
        ICell titleCell = titleRow.CreateCell(0);
        titleCell.SetCellValue("工　程　估　價　單");
        titleCell.CellStyle = styles.TitleStyle; // 應該已經有置中的 style
        sheet.AddMergedRegion(new CellRangeAddress(2, 2, 0, 6)); // A3~G3

        // ===== 工程名稱（A4~G4）=====
        IRow infoRow = sheet.CreateRow(rowIndex++);
        infoRow.HeightInPoints = 20;

        var projectCell = infoRow.CreateCell(0);
        projectCell.SetCellValue($"工程名稱：{request.ProjectName}");
        projectCell.CellStyle = styles.CellStyle; // 直接用靠左樣式
        sheet.AddMergedRegion(new CellRangeAddress(rowIndex - 1, rowIndex - 1, 0, 6)); // 合併該列 A~G

        // ===== 日期/計算/核算（A5~G5）=====
        IRow info2Row = sheet.CreateRow(rowIndex++);
        info2Row.HeightInPoints = 20;

        var dateCell = info2Row.CreateCell(0);
        dateCell.SetCellValue("日　　期：____________　計　　算：____________　核　　算：____________");
        dateCell.CellStyle = styles.CellStyle;
        sheet.AddMergedRegion(new CellRangeAddress(rowIndex - 1, rowIndex - 1, 0, 6)); // 合併該列 A~G

        // ===== 空一列 =====
        rowIndex++;

        // ===== 備註標題（A7~G7）=====
        IRow noteTitleRow = sheet.CreateRow(rowIndex++);
        var noteTitleCell = noteTitleRow.CreateCell(0);
        noteTitleCell.SetCellValue("備註：");
        noteTitleCell.CellStyle = styles.CellStyle;
        sheet.AddMergedRegion(new CellRangeAddress(rowIndex - 1, rowIndex - 1, 0, 6)); // 合併該列 A~G


        // ===== 備註內容 =====
        string[] noteContents = GetNoteContents(sheetType);

        for (int i = 0; i < noteContents.Length; i++)
        {
            IRow noteRow = sheet.CreateRow(rowIndex++);

            var numberCell = noteRow.CreateCell(0);
            numberCell.SetCellValue($"{i + 1}.");
            numberCell.CellStyle = styles.NoteCellStyle;

            var contentCell = noteRow.CreateCell(1);
            contentCell.SetCellValue(noteContents[i]);
            contentCell.CellStyle = styles.NoteCellStyle;
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex - 1, rowIndex - 1, 1, 6));

            double totalColWidth = 0;
            for (int col = 1; col <= 6; col++)
            {
                totalColWidth += sheet.GetColumnWidth(col);
            }
            int widthInChars = (int)(totalColWidth / 256);

            // 呼叫自適應列高的方法，字體大小可用樣式字體大小(12f)
            NpoiExcelUtility.AutoSizeRowHeight(noteRow, noteContents[i], widthInChars, 12f);
        }

    }

    // 設定備註頁面的欄位寬度
    private void SetNoteSheetColumnWidths(ISheet sheet)
    {
        sheet.SetColumnWidth(0, 5 * 256);      // A欄：編號
        sheet.SetColumnWidth(1, 80 * 256);     // B欄：內容（較寬）
        sheet.SetColumnWidth(2, 10 * 256);     // C欄
        sheet.SetColumnWidth(3, 10 * 256);     // D欄
        sheet.SetColumnWidth(4, 10 * 256);     // E欄
        sheet.SetColumnWidth(5, 10 * 256);     // F欄
        sheet.SetColumnWidth(6, 15 * 256);     // G欄
    }

    // 根據類型取得備註內容
    private string[] GetNoteContents(string sheetType)
    {
        if (sheetType == "結構體及外牆裝修")
        {
            return new string[]
             {
                "本工程估價單（全部）所列項目、數量(含損料)本於善良所計得之參考，投標廠商需確實探勘現場，並應詳閱招標文件圖說核算數量，若有項目或圖面未盡詳細或數量增減時，投標廠商得於投標報價前提出釋疑徵求院方或建築師說明後，投標廠商應將應作部份併入其他共同項目內或自行調整相關項目單價或自行填入第九項設計圖說內容或其他工程估價單未例入項目工料欄，投標廠商提出報價視為同意本工程(含各附件)所載，不得片面任意增減刪改工程估價單所列項目、數量、內容，暨於議價得標後承攬廠商須按契約圖說完成本工程且不得要求工程追加或藉故停工，否則須負賠償責任。",
                "本工程係總價承攬，承攬廠商不得因工資、物價、匯率、利率等之變動而要求增加任何費用。",
                "本工程估價單所列工程項目之單、複價不得填寫零數，惟經院方同意者除外，否則視為無效標。",
                "本工程估價單全份應填寫清楚並加蓋投標廠商及負責人印章、如有塗改應加蓋負責人印章。",
                "本工程估價單、施工補充說明書、施工規範附件、投標預知、廠商釋疑單及相關圖說等承攬廠商同意於締約時列為合約一部份。",
                "本工程位屬彰濱工業區管理局，所有工項工程(尤以模板工程、鷹架工程、鋼筋工程、臨時水電設施工程…等)須確實遵照管理局所頒布規章、規範並優先各項職業安全衛生法規之規定執行施工，若違反條文遭受罰款或要求停工時，承攬廠商無條件負全面責任。",
                "本工程承攬廠商須無條件配合(不因院方另包工程)自來水公司、電力公司、電信局、污水處理、消防局、電梯協會…等取得核准函並申請領得使用執照 並配合送水送電及取得營業許可，不得再另行追加費用。",
                "材料標單所列數量(含損料)僅供參考，投標廠商應依據設計圖面施工說明及親赴現場瞭解工地狀況後，做詳實精算，數量如有增減，應自行於標單最後第九項目，不得塗註本標單項目,廠牌規格，得標後視同全部估算完成，不得於工程進行中要求追加或藉詞推諉致延誤工期。",
                "本案所有在屋外外露部份及水池內及用水設施(含水管、冰水管、冷氣風管、泵浦…等)之另件、固定架、螺絲(室內外)等，均應為不鏽鋼製品。",
                "承包商應具備繪製施工圖之能力，並應指派具有經驗且有責任感之工地主任，於施工前應先將施工圖送監造單位審核，待核準後，方可按圖施工，否則有任何問題一切由承包商自行負責，並提出各階段工作進度表，材料設備進場與出工之配合彙總表。",
                "施工圖繪製費用含設計標準圖套繪成現場施工圖，各工種施工前應先提供施工圖給監造單位確認後，再行施工（包括磁磚及天花板分割計劃），甲方不負責施工圖正確性之責任。需配合客戶變更套繪施工圖，竣工圖修正提供藍曬圖及隨身碟等。",
                "窗戶底緣灌漿時內高外低，女兒牆外高內低。",
                "水電施作之穿樑套管，營造需加補強筋。",
                "浴室降板，另外浴室及管道間及濕室空間止水墩，應與板面結構一次完成，不得二次施工，且高於房間地板15公分。",
                "日用水池牆面及地面塗無毒防水層(出示證明)＋白色石英磚（含牆面）。",
                "空污費，保險費，檢據核銷，違規罰款，應由乙方自行負責。",
                "營建承包廠商借用電梯車廂施工期間施作保護層、保管、臨時電設置、保養費，施工中若有損壞，乙方應負賠償責任。",
                "營建承包廠商需負責電梯升降道內模板鐵絲螺絲及所有垃圾清除，爆模打除清運，機坑防水止漏責任施工。",
                "營建承包廠商應提供各樓層飾材完成面之一米線，若因澆置導致之高低落差超過正負3公分,應負責打除清運,不得衍生追加費用。",
                "規範、標單、圖面其中有提到之項目，乙方皆應施作，不得追加及推諉。",
                "如因甲方變更設計造成追加減帳，除新增品項須經甲乙雙方議價外，其餘皆以合約單價計。",
                "如防火門由業主自行發包，則防火門由得標廠商連工帶料施作，營造廠應負責結構M、O(門、窗與結構填塞沙漿)預留尺寸及配合施工廠商安裝的品質保證(高度要一致、完成面要注意)及收尾(矽利康、門框油漆、保護等)。",
                "如鋁窗框由業主自行發包，則鋁窗框由得標廠商連工帶料施作，營造廠應負責安裝的品質保證(窗戶的出入一致、高低一致)、漏水等問題的處理，及收尾(崁縫、塞水路、保護等)。",
                "電梯營造廠應負責借機時內裝保護及損壞賠償。",
                "欄杆、女兒牆、窗戶臺度尺寸，應以地坪飾材完成高程計算，符合建築法規規定淨高，不得以驅體結構尺寸及建築師圖面是結構尺寸，而衍生追加費用。",
                "玻璃才數計算方式：帷幕牆含框計算以90%計得，鋁窗含框計算以85%計得。",
                "本案如裝修影響工程進度，需配合展延，不另追加工程費用，含管理費、保險、勞安、保全…等。使用執照須含室內裝修申請",
                "防水保固 5 年。",
                "標單面積計算均已含損耗。",
                "完工時抽水井及觀測井預留各一處不拆除保留院方使用。",
                "汙水池採用無毒環氧樹脂施作。",
                "萬能鑰匙(Master Key)分配：本棟總樓層3支萬能鑰匙；總樓層機房、管道間3支萬能鑰匙；各樓層獨立各3支萬能鑰匙。",
                "觀測系統須有振動及位移觀測。",
                "截水溝單向洩水(1/100)。",
                "屋頂素地試水 24 小時，施作防水後 72 小時試水,材料為：PS 板+2500PSI+抗裂籤維，洩水坡度維 1/100~1/150 單面洩水。",
                "外牆面貼磁磚每 3M 打一個格子，0.3cm 水泥膠，打底 1cm。",
                "二丁掛磚完成須清洗完成。",
                "所有樓梯暨安全梯須前後錯階。",
                "樓梯扶手須固定預埋鐵件，扶手毛絲面不銹鋼厚度1.2mm,和接觸須 啞焊及拋光。",
                "外牆轉角均須做補強筋，開口處亦同。",
                "樓梯要施作預埋鋼筋及補強筋，樓板上層筋須放置樓梯下層搭接。",
                "蜂窩處理方式，高壓沖洗完後再以高強度水泥施作。",
                "預拌混凝土超過 90 分鐘(符合CNS規範標準)不得使用。",
                "各處露臺、屋頂女兒牆須留設溢水口。FL+15CM(完成面)",
                "補強筋的號數需大於主筋。例：主筋若 4 號，補牆筋 5 號。",
                "金屬門若有電動門需含重型門機、不銹鋼地軌及消防聯控設備及防夾、感應裝置、陽極鎖(金大漢)。",
                "本工程估價單備註暨施工補充說明書所列條文項目，所須費用均包含本工程內，承攬廠商不得以任何理由及其它因素要求增價及延長工期。",
                "合約單價，依議價降價比例調整。",
                "本工程速立康採用道康寧矽利康(型號： 室內潮濕防霉 991 / 818、 室外耐候氣密 791、高強度結構 995、天然石材 756 / 758)。",
                "專業技師辦理建築物結構與設備專業工程簽證費含於本工程合約內，費用由乙方支付費用 。",
                "合約甲乙雙方印花稅由乙方貼付足額，正本兩份，副本七份（甲方正本一份，副本七份，副本其中二份不須圖說，如乙方要保存副本自行增加）。",
                "全份合約(含標單、投標標價清單、單價分析、補充說明書、圖說、施工規範、履約保證書影本)製本由建築師事務所製作，承攬廠商支付費用。",
                "本次工程不含機電工程。",
                "全份合約(含圖說)製本由建築師事務所製作，承攬廠商支付費用。"
             };

        }
        else if (sheetType == "裝修")
        {
            return new string[]
            {
                "本工程估價單（全部）所列項目、數量本於善良所計得之參考，投標廠商需確實探勘現場，並應詳閱招標文件圖說核算數量，若有項目或圖面未盡詳細或數量增減時，投標廠商得於投標報價前提出釋疑徵求院方或建築師說明後，投標廠商應將應作部份併入其他共同項目內或自行調整相關項目單價或自行填入第十二項設計圖說內容或其他工程估價單未例入項目工料欄，投標廠商提出報價視為同意本工程(含各附件)所載，不得片面任意增減刪改工程估價單所列項目、數量、內容，暨於議價得標後承攬廠商須按契約圖說完成本工程且不得要求工程追加或藉故停工，否則須負賠償責任。",
                "本工程係總價承攬，承攬廠商不得因工資、物價、匯率、利率等之變動而要求增加任何費用。",
                "本工程估價單所列工程項目之單、複價不得填寫零數，惟經院方同意者除外，否則視為無效標。",
                "本工程估價單全份應填寫清楚並加蓋投標廠商及負責人印章、如有塗改應加蓋負責人印章。",
                "本工程估價單、施工補充說明書、施工規範附件、投標須知、廠商釋疑單及相關圖說等承攬廠商同意於締約時列為合約一部份。",
                "地下一層中央廚房範圍內所有室內裝修工程，需取得廚房設備得標廠商圖面簽認確定後方可施做。",
                "鉛鈑屏蔽及銅網牆面，於相關空間進場施作前，須先與該空間設備廠商確認進場設備並提送施作之各個輻防空間輻防檢測合格證明及輻防技師簽證，且須符合法規規定，經院方相關單位簽認後方可進行施作。",
                "本標案不允許廠商以電子檔出標。",
                "本工程位屬彰濱工業區管理局，所有工項工程(尤以施工架工程、臨時水電設施工程…等)須確實遵照管理局所頒布規章、規範並優先各項職業安全衛生法規之規定執行施工，若違反條文遭受罰款或要求停工時，承攬廠商無條件負全面責任。",
                "本工程包含自來水公司、電力公司、電信局及污水處理等等在本工程取得使用執照，送水送電及取得營業許可之間，申請手續及所有費用（含所有簽證），承包商需全權負責，不得再另行追加費用，惟行政規費，由業主負擔，乙方代墊(檢據核銷)。",
                "本次工程不含機電工程、結構體暨外飾門窗工程。",
                "承包商應具備繪製施工圖之能力，並應指派具有經驗且有責任感之工地主任，於施工前應先將施工圖送監造單位審核，待核準後，方可按圖施工，否則有任何問題一切由承包商自行負責，並提出各階段工作進度表，材料設備進場與出工之配合彙總表。",
                "施工圖繪製費用含設計標準圖套繪成現場施工圖，各工種施工前應先提供施工圖給監造單位確認後，再行施工（包括磁磚及天花板分割計劃），甲方不負責施工圖正確性之責任。需配合客戶變更套繪施工圖，竣工圖修正提供藍曬圖及隨身碟等。",
                "本案如有衛浴瓷器或主要器具主機由業主自行採購項目，承包商應負責連絡進場時程及貨到後之一切搬運、安裝、保管之責任。",
                "浴室排水：單向洩水1/100，檯面下靠牆明溝中間收方型不鏽鋼排水孔，淋浴間單向洩水1/100收至蓮蓬頭下靠牆全尺寸不鏽鋼水槽。",
                "衛浴設備由業主提供，乙方應於安裝前一個月提出出貨時程，貨至工地簽收後，乙方應負責保管，搬運，安裝之責任，至取得使用執照後驗收完成。",
                "發電機室、冰水主機等等重機械機房其隔音工程天花板亦需施作，應採浮動式吸震隔音基礎座。",
                "空污費檢據核銷，保險費，違規罰款，應由乙方自行負責。",
                "營造負責電梯車廂施工期間施作保護層、保管、臨時電設置、保養費，乙方施工如須使用，自行與營造協商共用，使用中若有損壞，乙方應負賠償責任。",
                "營造應提供各樓層飾材完成面之一米線，若因澆置導致之高低落差超過1公分,應負責打除清運,不得衍生追加費用。",
                "如因甲方變更設計造成追加減帳，除新增品項須經甲乙雙方議價外，其餘皆以合約單價計。",
                "防火門如有由業主自行發包，則防火門由得標廠商連工帶料施作，乙方應負責安裝的品質保證(高度要一致、完成面要注意)及收尾崁縫(矽利康、門框油漆、保護等)。",
                "鋁窗如有由業主自行發包，則鋁窗由得標廠商連工帶料施作，乙方應負責安裝的品質保證(窗戶的出入一致、高低一致)、漏水等問題的處理，及收尾(崁縫、矽利康、塞水路、保護等)。",
                "電梯如有由業主自行發包，則得標廠商連工帶料施作，乙方應負責安裝的品質保證、門框崁縫、借機時內裝保護等。",
                "天花板輕鋼架應使用2分螺桿吊筋。廚房用2分不鏽鋼螺桿吊筋。",
                "防撞安全扶手、廁所浴室搗擺及門片(含圖視、掛勾)、小便斗搗擺、洗手檯、無障礙廁所金屬扶手、嬰兒換尿布檯、固定式嬰兒安全座椅、置物架……等等所有固定於牆面之物件，應於輕隔間內補強鐵片厚度至少1.25mm。",
                "暗架矽酸鈣板，應留60*60cm維修孔(含機電、空調設備位置留設)。",
                "欄杆、女兒牆、窗戶抬度等尺寸，完工後，應符合建築法規規定，不得因建築師圖面是結構尺寸而衍生追加。",
                "玻璃才數計算方式：帷幕牆含框計算以90%計得，鋁窗含框計算以85%計得。",
                "本案如結構體暨外飾門窗、機電及其他……影響工程進度，需配合展延，不另追加工程費用，含管理費、保險、勞安、保全…等。使用執照申領時室內裝修承攬廠商須配合結構體暨外飾門窗工程廠商，申領使用執照及開業申請間，應提供申請所需各項證明文件及申請書必要之用印和簽証。",
                "室內裝修工程保固 ２ 年。電動門機(捲門、橫拉門)保養一年一次，共二次",
                "工程估價單數量本於善良計得，耗損計入單價內。",
                "確認得標之廠商，在雙方完成契約用印之前，若工地現場需要協調界面或確認施工細節時，須配合需求於通知協調日期進場協商。",
                "萬能鑰匙(Master Key)分配：本棟總樓層3支萬能鑰匙；總樓層機房、管道間一支萬能鑰匙；各樓層獨立各3支萬能鑰匙(結合結構體暨外飾門窗工程配設萬能鑰匙)。",
                "矽酸鈣板明架天花完工後預留使用總數量1%數量無償供院方做備品。",
                "截水溝單向洩水(1/100)，兩側均須設置排水口,採 3\"PVC，排水管配置不包含於本裝修工程，截水溝排水孔開孔2\"安裝。",
                "櫥櫃內外部或檯面、水槽檯面，在機電工程、資訊工程或其它工程內容須設插座、開關、給水龍頭、排水管路或其他須穿孔留洞時，均無條件配合引孔留設，室內裝修承攬廠商不含水電、弱電及消防配管線工程。",
                "承攬廠商須按承包金額千分之一.五費用提撥當工地專用清安基金。(由工地清安協議組織統籌管理應用並採多退少補。)",
                "天花、牆面、地坪交屋前須清洗或擦拭完成後點交。AB膠+補土+研磨+底漆+面漆",
                "樓梯扶手須固定預埋鐵件，扶手毛絲面不銹鋼厚度 1.2mm,和接觸須 啞焊及拋光。",
                "室內輕隔間 2 塊板片要交丁施做，不可對縫，油漆處接縫施作須以AB膠 +補土+不織布+研磨拋光+ 底漆+面漆。",
                "門邊輕隔間板，應以菜刀板工法施作。",
                "金屬門若為電動門需含重型門機(金大漢)、不銹鋼地軌及消防聯控設備及防夾、感應裝置，管路及配線乾接點及全組設備齊全，不含門禁系統、消防磁力門扣及其配管配線。",
                "本工程貼飾磁磚全採送審核准磁磚專用黏著劑，不得採用海菜粉。",
                "輕隔間工程須出具一小時防火時效認可通知書，並施作一小時防火時效牆認可通知書內規範之填縫與收邊；矽酸鈣板及纖維水泥板須出具耐燃一級登錄證書。",
                "天花板材須為耐燃一級，且有經濟部商品檢驗局商品驗證登錄証書。",
                "防火門須附防火證明，並符合消防法規及認証，且可取得使用執照。",
                "系統櫃桶身及底板六分木心板，內側貼Melamine paper，外側可見面貼Melamine paper1.0mm厚富美家美耐板，門板木心板六分粒片板(潤濕抗彎試驗≧9 N/mm)雙面貼Melamine paper1.0mm厚富美家美耐板，彎曲檯面十分木心板面貼富美家彎曲美耐板一體成形，及天花板及固定端牆面收邊板材，板材均為1mm ABS封邊。鉸鏈採國產金屬110°緩衝鉸鏈。踢腳板採六分木心板面貼Melamine paper粒片板面貼1.0mm厚富美家美耐板(潤濕抗彎試驗≧8.5 N/mm)。抽屜用金屬三截式滾珠緩衝滑軌。櫥櫃鎖匙每單位配3支Master key，更衣櫃需可使用員工識別證刷卡的電子鎖。美耐板、五金配件全送樣另選，無防火及防焰證明。",
                "本木作系統櫃均依單位須求施作鎖頭，並製作鎖王系統，鎖王系統的各組範圍依使用單位須求製作，施作前廠商應與使用單位確認。",
                "系統櫃圖說標示之材料及配件，五金另料屬本次發包內容均須施作。",
                "斜把手後面補板(門片、抽屜)屬系統櫃工程，不另加價。",
                "地坪飾材以結構體營造廠所提供各樓層飾材完成面之1米線為基準，以紅外線雷射儀整平1:3水泥砂漿地坪（磁磚、石材地坪）及1:3水泥砂漿粉光地坪（磁磚、石材地坪以外），@2M*2M平整度±5MM內。地坪磁磚、石材飾材施作前須先檢送施工計畫說明濕式或乾式施工及整平固定器布置，且飾材完成面與周遭地坪飾材完成面平順無高低差，並計算磁磚厚度及洩水坡厚度，自行計減地坪應填厚度，經核准後方可進場施工。",
                "如牆面地坪採用之磁磚及人造石材由業主自行採購，室內裝修廠商須連工帶料（底材料）施工並負責面貼平整及品質保證和收尾抹縫。",
                "如地坪無縫PVC地毯由業主自行發包，室內裝修廠商須將地坪施工完整交予無縫PVC地毯廠商後續進度，若有工程瑕疵，室內裝修廠商無條件改善。",
                "門框、窗框安裝後四周與牆介面由室內裝修廠商(本工程)施作防霉矽利康，以及崁縫塞水路收尾。",
                "地坪1:3水泥砂漿粉刷或粉光及板降區回填混凝土，須依結構體廠商各樓層現場所做飾材完成面1m基準線，採紅外線雷射儀器精準水平放樣，@2M*2M平整度±5MM內，清除粉塵噴塗水泥介面劑再施作所需工項，並精算預留地坪飾材施作厚度，完成後地坪須平順無高低差。",
                "B1F中央廚房範圍各項室內裝修工程，需由廚房設備工程得標廠商簽認及現場放樣後，方可施作。",
                "本工程速立康採用道康寧矽利康(型號： 室內潮濕防霉 991 / 818、 室外耐候氣密 791、高強度結構 995、天然石材 756 / 758)。",
                "空污費，保險費，檢據核銷，違規罰款，應由乙方自行負責。",
                "本工程含五大管線專業技師簽證費。",
                "本工程估價單、 估價單備註、標單、圖說、施工規範暨施工補充說明書所列條文項目，乙方皆應施作，所須費用均包含本工程內，承攬廠商不得以任何理由及其它因素要求增價及延長工期。",
                "合約單價，依議價報價單降價比例調整。",
                "全份合約(含標單、投標標價清單、單價分析、補充說明書、圖說、施工規範、履約保證書影本)製本由建築師事務所製作，承攬廠商支付費用。",
                "合約甲乙雙方印花稅由乙方貼付足額，正本兩份，副本七份（甲方正本一份，副本七份，副本其中二份不須圖說，如乙方要保存副本自行增加）。",
                "全份合約(含圖說)製本由建築師事務所製作，承攬廠商支付費用。"
            };
        }

        return new string[] { "暫無備註內容" };
    }
}