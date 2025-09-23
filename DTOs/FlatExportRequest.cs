namespace ShowPms.DTOs
{
    public class FlatExportRequest
    {
        public string ProjectName { get; set; }
        public List<FlatExportRequestItem> Items { get; set; }
    }

    public class FlatExportRequestItem
    {
        public int Id { get; set; }
        public string Code { get; set; }
        public string Vender { get; set; }
        public string Name { get; set; }
        public string Spec { get; set; }
        public string Unit { get; set; }
        public string Quantity { get; set; }
        public string UnitPrice { get; set; }
        public string ContractUnitPrice { get; set; }
        public string Note { get; set; }
        public string MajorCategory { get; set; }
        public string MiddleCategory { get; set; }
    }
}
