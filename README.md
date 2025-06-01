# PdfExcelReportTool

**PdfExcelReportTool** is a .NET-based reporting solution that generates reports in **PDF** and **Excel** formats using two powerful libraries: [QuestPDF](https://www.questpdf.com/) for PDF generation and [ClosedXML](https://github.com/ClosedXML/ClosedXML) for Excel.

The reports are based on a structured data model and maintain the same layout and content across both formats. This consistency ensures a uniform presentation regardless of the chosen output format.

## 📄 Features

- Generate stylized PDF reports using **QuestPDF**
- Generate Excel reports using a **template-based approach** with **ClosedXML**
- Maintain consistent formatting across both formats
- Populate data dynamically from a source list
- Automatically replicate rows while keeping the original cell styling in Excel

## 🛠 Technologies Used

- [.NET 9](https://dotnet.microsoft.com/)
- [QuestPDF](https://github.com/QuestPDF/QuestPDF) – for PDF generation
- [ClosedXML](https://github.com/ClosedXML/ClosedXML) – for Excel generation

## 🚀 How It Works

### ✅ PDF Report Generation (QuestPDF)

The report is built using QuestPDF's fluent API. Before generating any documents, you must **define the license** type to comply with QuestPDF’s terms:

```csharp
QuestPDF.Settings.License = LicenseType.Community;
```

### ✅ Excel Report Generation (ClosedXML)
For Excel, the project uses a template-based technique:

1. A preformatted Excel template is created with headers and styles.
2. The code imports this template.
3. The first data row is cloned and reused to maintain styling.
4. Each subsequent record is added by duplicating the formatted row, preserving design consistency.

This approach ensures the Excel file mirrors the PDF layout, including fonts, borders, colors, and alignments.
