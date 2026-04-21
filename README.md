# OpenXMLSDK.Engine

A .NET library that simplifies generating and manipulating Word (.docx) documents using the OpenXML SDK. Use the `WordManager` object and the built-in Report Engine to produce rich documents from templates or from scratch.


## Quality and packaging

[![Quality Gate Status](https://sonarcloud.io/api/project_badges/measure?project=mathieumack_OpenXMLSDK.Engine&metric=alert_status)](https://sonarcloud.io/summary/new_code?id=mathieumack_OpenXMLSDK.Engine)
[![.NET](https://github.com/mathieumack/OpenXMLSDK.Engine/actions/workflows/ci.yml/badge.svg)](https://github.com/mathieumack/OpenXMLSDK.Engine/actions/workflows/ci.yml)
[![NuGet package](https://buildstats.info/nuget/OpenXMLSDK.Engine?includePreReleases=true)](https://nuget.org/packages/OpenXMLSDK.Engine)


## API

The API of `WordManager` is designed to be easy to understand and use.

### Open an existing template

Call `OpenDocFromTemplate` to open a `.dotx` template, copy it to a new path and work on the copy:

```csharp
var templatePath = @"C:\temp\template.dotx";
var outputPath   = @"C:\temp\createdDoc.docx";

using (var word = new WordManager())
{
    word.OpenDocFromTemplate(templatePath, outputPath, isEditable: true);

    // ... make changes ...

    word.SaveDoc();
    word.CloseDoc();
}
```

You can also open a template from a `Stream` (useful in web or cloud scenarios):

```csharp
using (var templateStream = File.OpenRead(@"C:\temp\template.dotx"))
using (var word = new WordManager())
{
    word.OpenDocFromTemplate(templateStream);

    // ... make changes ...

    word.SaveDoc();
    var result = word.GetMemoryStream(); // returns a copy of the document as a MemoryStream
}
```

### Insert text on a bookmark

Templates can contain named bookmarks. Use the bookmark extension methods to inject content:

```csharp
using (var word = new WordManager())
{
    word.OpenDocFromTemplate(templatePath, outputPath, isEditable: true);

    // Insert a plain string at a bookmark named "CompanyName"
    word.SetTextOnBookmark("CompanyName", "Contoso Ltd.");

    // Insert multiple lines at a bookmark named "AddressLines"
    word.SetTextsOnBookmark("AddressLines", new List<string>
    {
        "1 Microsoft Way",
        "Redmond, WA 98052"
    });

    // Replace a bookmark with HTML content
    word.SetHtmlOnBookmark("BodyContent", "<p>Hello <strong>World</strong></p>");

    word.SaveDoc();
    word.CloseDoc();
}
```

### Report Engine – generate a document from a model

The Report Engine lets you define an entire document as a C# object tree, bind it to a `ContextModel` and render it to bytes in one call.

#### Build a simple document

```csharp
using OpenXMLSDK.Engine.Word;
using OpenXMLSDK.Engine.Word.ReportEngine;
using OpenXMLSDK.Engine.Word.ReportEngine.Models;
using OpenXMLSDK.Engine.ReportEngine.DataContext;
using OpenXMLSDK.Engine.ReportEngine.DataContext.FluentExtensions;
using System.Globalization;

// 1. Build the document model
var document = new Document();
var page = new Page();
document.Pages.Add(page);

var paragraph = new Paragraph();
page.ChildElements.Add(paragraph);

paragraph.ChildElements.Add(new Label { Text = "Hello, #CustomerName#!" });

// 2. Build the context (data bindings)
var context = new ContextModel()
    .AddString("#CustomerName#", "Alice");

// 3. Generate bytes
using var word = new WordManager();
byte[] docBytes = word.GenerateReport(document, context, CultureInfo.InvariantCulture);
File.WriteAllBytes(@"C:\temp\output.docx", docBytes);
```

#### Render a table from a collection

```csharp
var document = new Document();
var page = new Page();
document.Pages.Add(page);

var table = new Table
{
    DataSourceKey = "#Orders#",
    ColsWidth = new[] { 3000, 3000, 2000 },
    HeaderRow = new Row
    {
        Cells =
        {
            new Cell { ChildElements = { new Label { Text = "Product",  Bold = true } } },
            new Cell { ChildElements = { new Label { Text = "Quantity", Bold = true } } },
            new Cell { ChildElements = { new Label { Text = "Price",    Bold = true } } },
        }
    },
    RowModel = new Row
    {
        Cells =
        {
            new Cell { ChildElements = { new Label { Text = "#Product#"  } } },
            new Cell { ChildElements = { new Label { Text = "#Quantity#" } } },
            new Cell { ChildElements = { new Label { Text = "#Price#"    } } },
        }
    }
};
page.ChildElements.Add(table);

// Build a data-source collection
var context = new ContextModel()
    .AddCollection("#Orders#",
        new ContextModel().AddString("#Product#", "Widget A").AddString("#Quantity#", "10").AddString("#Price#", "$5.00"),
        new ContextModel().AddString("#Product#", "Widget B").AddString("#Quantity#", "3").AddString("#Price#", "$15.00")
    );

using var word = new WordManager();
byte[] docBytes = word.GenerateReport(document, context, CultureInfo.InvariantCulture);
```

#### Multi-report generation (multiple sections)

```csharp
var reports = new List<Report>
{
    new Report
    {
        Document     = BuildCoverPageDocument(),
        ContextModel = BuildCoverPageContext(),
        AddPageBreak = true
    },
    new Report
    {
        Document     = BuildContentDocument(),
        ContextModel = BuildContentContext(),
        AddPageBreak = false
    }
};

using var word = new WordManager();
byte[] docBytes = word.GenerateReport(reports, mergeStyles: true, CultureInfo.CurrentCulture);
```

### Append an external document

Append one or more Word documents at the end of the current document:

```csharp
using (var word = new WordManager())
{
    word.OpenDocFromTemplate(templatePath, outputPath, isEditable: true);

    using var extra = File.OpenRead(@"C:\temp\annex.docx");
    word.AppendSubDocument(extra, withPageBreak: true);

    word.SaveDoc();
    word.CloseDoc();
}
```

### Create a new blank document

```csharp
using (var word = new WordManager())
{
    word.New();

    // ... add content programmatically ...

    word.SaveDoc();
    var stream = word.GetMemoryStream();
}
```

### Context fluent extensions reference

`ContextModel` supports a fluent API for adding typed values:

| Method | Description |
|---|---|
| `AddString(key, value)` | Plain string |
| `AddDouble(key, value, pattern)` | Formatted double (e.g. `"{0:N2}"`) |
| `AddBoolean(key, value)` | Boolean (controls `Show`, `Bold`, etc.) |
| `AddDateTime(key, value, pattern)` | Formatted date/time |
| `AddByteContent(key, bytes)` | Inline image from a byte array |
| `AddBase64Content(key, base64)` | Inline image from a Base64 string |
| `AddFileLink(key, filePath)` | File reference (image path or any file path used by the report engine) |
| `AddSubstitutableString(key, pattern, dataSource)` | Composite string, e.g. `"{0} of {1}"` |
| `AddCollection(key, contexts...)` | Data source for `ForEach` or table row models |


# Contribute

## How to contribute

If you want to contribute to this project, you can do it in several ways:

- [Submit bugs and feature requests]
- [Review source code changes]
- [Review the documentation and make pull requests for anything from typos to new content]