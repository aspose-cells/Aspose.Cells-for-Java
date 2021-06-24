![GitHub release (latest by date)](https://img.shields.io/github/v/release/aspose-cells-cloud/aspose-cells-cloud-java) ![GitHub all releases](https://img.shields.io/github/downloads/aspose-cells/Aspose.cells-for-Java/total) ![GitHub](https://img.shields.io/github/license/aspose-cells/Aspose.cells-for-java)
# Java API for Excel File Formats

[Aspose.Cells for Java](https://products.aspose.com/cells/java) is an award-winning Excel Spreadsheet Processing API that allows Java developers to embed the ability to read, write and manipulate Excel® spreadsheets (XLS, XLSX, XLSM, XLSB, XLTX, SpreadsheetML, CSV, ODS), HTML, MHTML, PDF, and image file formats into their own Java applications without needing to rely on Microsoft Excel®.

Directory | Description
--------- | -----------
[Examples](https://github.com/aspose-cells/Aspose.Cells-for-Java/tree/master/Examples) | A collection of Java examples that help you learn the product features.
[Examples.GridWeb](https://github.com/aspose-cells/Aspose.Cells-for-Java/tree/master/Examples.GridWeb) | A collection of Java examples that help you learn and explore Aspose.GridWeb features.
[Plugins](https://github.com/aspose-cells/Aspose.Cells-for-Java/tree/master/Plugins) | Plugins that will demonstrate one or more features of Aspose.Cells for Java.

<p align="center">
  <a title="Download ZIP" href="https://github.com/aspose-cells/Aspose.Cells-for-Java/archive/master.zip">
    <img src="https://raw.githubusercontent.com/AsposeExamples/java-examples-dashboard/master/images/downloadZip-Button-Large.png" alt="Download Aspose.Cells for Java Examples, Plugins and Showcases" />
  </a>
</p>

## Excel File Processing Features

### Document Features

- Open Plain or Encrypted Excel files (Excel97, Excel2007/2010/2013) from different sources.
- Save Excel files (Excel97- Excel2007/2010/2013) in various supported formats.
- Convert Excel files & spreadsheets to various supported formats.
- Convert to Tagged Image File Format (`TIFF`).
- Read and Write OpenDocument Spreadsheet (`ODS`) format.
- Modify the document properties of Excel files.

### Worksheet Features

- Make Worksheet visible or hidden.
- Ability to show or hide worksheet tabs, scroll bars, gridlines & headers.
- Apply worksheet zoom level.
- Keep the selected data visible while scrolling in freeze panes.
- Ability to preview worksheet page breaks.
- Protection support for worksheet content, objects as well as scenarios.
- Perform and apply page setup configuration to worksheets.
- Perform various actions on individual or group of rows and columns.

### Data Management Features

- Insert data in specific cells at runtime.
- Fetch data from various data soures and import into worksheets.
- Retrieve data from cells based on their datatype.
- Get data from worksheet cells and export to array.
- Apply conditional formatting.
- Perform numerous formatting actions on data, such as, font setting.

### Charting & Graphics Features

- Supports creating various kinds of charts.
- Add custom charts to the worksheet.
- Add pictures to worksheets at the runtime.
- Ability to print worksheets.

### Advanced Features

- Use robust Formula Calculation Engine to support formula calculation.
- Manipulate VBA code or Macros.
- Create pivot tables as well as change its source data at runtime.

## Read & Write Spreadsheet Formats

**Microsoft Excel:** XLS, XLSX, XLSB, XLT, XLTX, XLTM, XLSM, XML\
**OpenOffice:** ODS\
**Text:** CSV, TSV\
**Web:** HTML, MHTML\
**Numbers:** Apple's iWork office suite Numbers app documents

## Save Excel Files As

**Fixed Layout:** PDF, PDF/A, XPS\
**Data Interchange:** DIF\
**Images:** JPEG, PNG, BMP, SVG, TIFF, EMF, GIF

## Supported Environments

- **Microsoft Windows:** Windows Desktop & Server (x86, x64)
- **macOS:** Mac OS X
- **Linux:** Ubuntu, OpenSUSE, CentOS, and others
- **Java Versions:** `J2SE 7.0 (1.7)`, or above

## Get Started with Aspose.Cells for Java

Aspose hosts all Java APIs at the [Aspose Repository](https://repository.aspose.com/webapp/#/artifacts/browse/tree/General/repo/com/aspose/aspose-cells). You can easily use Aspose.Cells for Java API directly in your Maven projects with simple configurations. For the detailed instructions please visit [Installing Aspose.Cells for Java from Maven Repository](https://docs.aspose.com/cells/java/installation/) documentation page.

## Convert Table to Range with Options using Java

```java
// For complete examples and data files, please go to https://github.com/aspose-cells/Aspose.Cells-for-Java
// The path to the documents directory.
String dataDir = Utils.getSharedDataDir(ConvertTableToRangeWithOptions.class) + "Tables/";
// Open an existing file that contains a table/list object in it
Workbook workbook = new Workbook(dataDir + "book1.xlsx");

TableToRangeOptions options = new TableToRangeOptions();
options.setLastRow(5);

// Convert the first table/list object (from the first worksheet) to normal range
workbook.getWorksheets().get(0).getListObjects().get(0).convertToRange(options);

// Save the file
workbook.save(dataDir + "ConvertTableToRangeWithOptions_out.xlsx");
```

[Home](https://www.aspose.com/) | [Product Page](https://products.aspose.com/cells/java) | [Docs](https://docs.aspose.com/cells/java/) | [Demos](https://products.aspose.app/cells/family) | [API Reference](https://apireference.aspose.com/cells/java) | [Examples](https://github.com/aspose-cells/Aspose.Cells-for-Java) | [Blog](https://blog.aspose.com/category/cells/) | [Search](https://search.aspose.com/) | [Free Support](https://forum.aspose.com/c/cells) | [Temporary License](https://purchase.aspose.com/temporary-license)
