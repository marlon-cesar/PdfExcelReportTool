using ClosedXML.Excel;
using PdfExcelReportTool.Shared;

var reportTemplate = Path.Combine(Directory.GetCurrentDirectory(), "ReportTemplate", "OrderReportTemplate.xlsx");

var workbook = new XLWorkbook(reportTemplate);
var worksheet = workbook.Worksheets.First();

var row = worksheet.LastRowUsed()!.RowNumber();
var firstRow = row;


var orders = OrderGenerator.GenerateOrders();

foreach (var order in orders)
{
    // Don't duplicate the first row
    if (firstRow != row)
        DuplicateRow(worksheet, firstRow, row);

    worksheet.Cell(row, 1).Value = order.Id;
    worksheet.Cell(row, 2).Value = order.CustomerName;
    worksheet.Cell(row, 3).Value = order.ProductName;
    worksheet.Cell(row, 4).Value = order.Quantity;
    worksheet.Cell(row, 5).Value = order.UnitPrice;
    worksheet.Cell(row, 6).Value = order.TotalPrice;
    worksheet.Cell(row, 7).Value = order.OrderDate;
    worksheet.Cell(row, 8).Value = order.Status;

    row++;
}



workbook.SaveAs(@$"c:\temp\report-{DateTime.Now:yyyy-MM-dd-HHmmss}.xlsx");



void DuplicateRow(IXLWorksheet worksheet, int rowFrom, int rowTo) =>
    worksheet.Row(rowFrom).CopyTo(worksheet.Row(rowTo));
