namespace ShowPms.DTOs
{
    public class OldPriceItemDto
    {
        public long Id { get; set; }
        public string Vender { get; set; }
        public string Name { get; set; }
        public string Unit { get; set; }
        public decimal? Quantity { get; set; }
        public decimal? UnitPrice { get; set; }
        public decimal? Amount { get; set; }
        public decimal? ContractUnitPrice { get; set; }
        public decimal? ContractAmount { get; set; }
        public string Note { get; set; }

        public string MinorCategoryName { get; set; }
        public string MiddleCategoryName { get; set; }
        public string MajorCategoryName { get; set; }
    }
}
