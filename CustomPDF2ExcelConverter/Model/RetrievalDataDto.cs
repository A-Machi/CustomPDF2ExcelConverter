namespace CustomPDF2ExcelConverter.Model
{
    public class RetrievalDataDto
    {
        public uint RowNumber { get; set; } = 0;

        public string OrderNumber { get; set; } = string.Empty;

        public string Plant { get; set; } = string.Empty;

        public string UnloadingPoint { get; set; } = string.Empty;

        public string ItemNumberCustomer { get; set; } = string.Empty;

        public string Naming { get; set; } = string.Empty;

        public string Status { get; set; } = "NEU";

        public string LastDelivery { get; set; } = string.Empty;

        public string WECaptureDate { get; set; } = string.Empty;

        public string Appointment { get; set; } = string.Empty;

        public string Quantity { get; set; } = "0";

        public string QuantityChange { get; set; } = "0";

        public string PTSOrder { get; set; } = "NEU";

        public string VNumber { get; set; } = string.Empty;
        
        public string Changes { get; set; } = string.Empty;
        
        public string Link { get; set; } = string.Empty;

        public string Remark { get; set; } = string.Empty;

        public string PriceBetween1_4 { get; set; } = string.Empty;

        public string PriceBetween5_9 { get; set; } = string.Empty;

        public string PriceBetween10_24 { get; set; } = string.Empty;

        public string PriceBetween25_49 { get; set; } = string.Empty;
        
        public string PriceBetween50_99 { get; set; } = string.Empty;

        public string PriceUpTo100 { get; set; } = string.Empty;

        public string VKValidSince { get; set; } = string.Empty;

        public string EKDelivery { get; set; } = string.Empty;

        public string EKDeliverySince { get; set; } = string.Empty;

        public string MCertificate { get; set; } = string.Empty;

    }
}