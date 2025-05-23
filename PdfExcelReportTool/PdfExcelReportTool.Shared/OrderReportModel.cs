namespace PdfExcelReportTool.Shared
{
    public record class OrderReportModel(
        int Id,
        string CustomerName,
        DateTime OrderDate,
        string ProductName,
        int Quantity,
        decimal UnitPrice,
        string Status
    )
    {
        public decimal TotalPrice => Quantity * UnitPrice;
    }

}
