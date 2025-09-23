namespace ShowPms.DTOs
{
    public class ExportRequest
    {
        public string ProjectName { get; set; }
        public List<MajorCategoryDto> MajorItems { get; set; }
    }

    public class MajorCategoryDto
    {
        public string Name { get; set; }   // e.g. "建築工程"
        public List<MiddleCategoryDto> MiddleItems { get; set; }
    }

    public class MiddleCategoryDto
    {
        public string Name { get; set; }  // e.g. "混凝土"
        public List<ExportRequestItem> Items { get; set; }
    }

    public class ExportRequestItem
    {
        public int Id { get; set; }
        public string Code { get; set; }
        public string Vender { get; set; }
        public string Name { get; set; }
        public string Spec { get; set; }
        public string Unit { get; set; }

        public string Quantity { get; set; } = "0";
        public string UnitPrice { get; set; } = "0";
        public string ContractUnitPrice { get; set; } = "0";
        public string Note { get; set; } = "";

        public decimal QuantityDecimal => decimal.TryParse(Quantity, out var vq) ? vq : 0;
        public decimal UnitPriceDecimal => decimal.TryParse(UnitPrice, out var vu) ? vu : 0;
        public decimal ContractUnitPriceDecimal => decimal.TryParse(ContractUnitPrice, out var vc) ? vc : 0;

        public decimal GetTotalPrice()
        {
            return QuantityDecimal * UnitPriceDecimal;
        }

        public decimal GetContractTotalPrice()
        {
            return QuantityDecimal * ContractUnitPriceDecimal;
        }
    }


}
