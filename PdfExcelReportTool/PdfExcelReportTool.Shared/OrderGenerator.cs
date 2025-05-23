namespace PdfExcelReportTool.Shared
{
    public static class OrderGenerator
    {
        private static readonly string[] CustomerNames = { "Alice", "Bob", "Charlie", "Diana", "Eve", "Frank", "Grace", "Henry" };
        private static readonly string[] ProductNames = { "Laptop", "Phone", "Tablet", "Monitor", "Keyboard", "Mouse", "Printer" };
        private static readonly string[] Statuses = { "Pending", "Completed", "Cancelled" };

        public static List<OrderReportModel> GenerateOrders(int count = 20)
        {
            var orders = new List<OrderReportModel>();
            var rand = new Random();

            for (int i = 1; i <= count; i++)
            {
                var customer = CustomerNames[rand.Next(CustomerNames.Length)];
                var product = ProductNames[rand.Next(ProductNames.Length)];
                var quantity = rand.Next(1, 10);
                var unitPrice = Math.Round((decimal)(rand.NextDouble() * 1000 + 50), 2); // $50 to $1050
                var status = Statuses[rand.Next(Statuses.Length)];
                var orderDate = DateTime.Today.AddDays(-rand.Next(0, 90)); // Last 90 days

                var order = new OrderReportModel(
                    Id: i,
                    CustomerName: customer,
                    OrderDate: orderDate,
                    ProductName: product,
                    Quantity: quantity,
                    UnitPrice: unitPrice,
                    Status: status
                );

                orders.Add(order);
            }

            return orders;
        }
    }

}
