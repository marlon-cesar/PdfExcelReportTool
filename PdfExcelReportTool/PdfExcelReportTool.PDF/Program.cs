using PdfExcelReportTool.Shared;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

QuestPDF.Settings.License = LicenseType.Community;
QuestPDF.Settings.CheckIfAllTextGlyphsAreAvailable = false;


var pdf = Document.Create(document =>
{
    _ = document.Page(page =>
    {
        CreatePage(page, "Order Report", false);

        CreateContent(page);
    });
});

pdf.GeneratePdf(@$"c:\temp\report-{DateTime.Now:yyyy-MM-dd-HHmmss}.pdf");
//var byteArray = pdf.GeneratePdf();




static PageDescriptor CreatePage(PageDescriptor page, string title, bool landscape)
{
    // Page info
    page.Size(landscape ? PageSizes.A4.Landscape() : PageSizes.A4.Portrait());
    page.Margin(20);
    page.PageColor(Colors.White);

    // Header
    page.Header()
                .Height(80)
                .ExtendHorizontal()
                .AlignCenter()
                .AlignTop()
                .Element(container => CreateHeader(container, title));

    // Footer
    page.Footer()
            .AlignCenter()
            .AlignMiddle()
            .Text(x => x.CurrentPageNumber());

    return page;
}

static void CreateHeader(IContainer doc, string title)
{
    var blastoiseImage = Path.Combine(Directory.GetCurrentDirectory(), "images", "blastoise.png");
    var charizardImage = Path.Combine(Directory.GetCurrentDirectory(), "images", "charizard.png");

    doc.Column(column =>
        column.Item().Table(table =>
        {
            table.ColumnsDefinition(columns =>
            {
                columns.ConstantColumn(100);
                columns.RelativeColumn();
                columns.ConstantColumn(100);
            });

            table.Cell().Column(c => c.Item().Height(70).Width(70).AlignLeft().AlignMiddle().Image(blastoiseImage).FitArea());
            table.Cell().ScaleToFit().Column(c => c.Item().Height(70).AlignCenter().AlignMiddle().Text(title).Style(TextStyle.Default.FontSize(16).Bold()));
            table.Cell().Column(c => c.Item().Height(70).Width(70).AlignRight().AlignMiddle().Image(charizardImage).FitArea());
        })
    );
}

static void CreateContent(PageDescriptor page)
{
    var orders = OrderGenerator.GenerateOrders();

    page.Content().Column(column =>
    {
        column.Spacing(10);

        column.Item().Table(table =>
        {
            table.ColumnsDefinition(colunas =>
            {
                colunas.ConstantColumn(50);
                colunas.RelativeColumn();
                colunas.RelativeColumn();
                colunas.RelativeColumn();
                colunas.RelativeColumn();
                colunas.RelativeColumn();
                colunas.RelativeColumn();
                colunas.RelativeColumn();
            });

            // header
            table.Header(header =>
            {
                header.Cell().Element(e => HeaderStyle(e)).Text("Id");
                header.Cell().Element(e => HeaderStyle(e)).Text("Customer");
                header.Cell().Element(e => HeaderStyle(e)).Text("Product");
                header.Cell().Element(e => HeaderStyle(e)).Text("Quantity");
                header.Cell().Element(e => HeaderStyle(e)).Text("Price");
                header.Cell().Element(e => HeaderStyle(e)).Text("Total");
                header.Cell().Element(e => HeaderStyle(e)).Text("Date");
                header.Cell().Element(e => HeaderStyle(e)).Text("Status");
            });

            foreach (var order in orders)
            {
                table.Cell().Element(e => DefaultCellStyle(e)).Text(order.Id.ToString());
                table.Cell().Element(e => DefaultCellStyle(e)).Text(order.CustomerName);
                table.Cell().Element(e => DefaultCellStyle(e)).Text(order.ProductName);
                table.Cell().Element(e => DefaultCellStyle(e)).Text(order.Quantity.ToString());
                table.Cell().Element(e => DefaultCellStyle(e)).Text(order.UnitPrice.ToString("C2"));
                table.Cell().Element(e => DefaultCellStyle(e)).Text(order.TotalPrice.ToString("C2"));
                table.Cell().Element(e => DefaultCellStyle(e)).Text(order.OrderDate.ToShortDateString());
                table.Cell().Element(e => DefaultCellStyle(e)).Text(order.Status);
            }

        });
    });
}

static IContainer HeaderStyle(IContainer elemento) =>
    elemento.Border(1).BorderColor(Colors.Black).Background(Colors.Yellow.Medium).AlignCenter().AlignMiddle().Padding(2).DefaultTextStyle(style => style.Bold().FontSize(9));

static IContainer DefaultCellStyle(IContainer elemento) =>
    elemento.Border(1).BorderColor(Colors.Black).Background(Colors.White).AlignCenter().AlignMiddle().Padding(2).DefaultTextStyle(style => style.FontSize(9).FontColor(Colors.Black));

