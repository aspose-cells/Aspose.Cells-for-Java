// For complete examples and data files, please go to https://github.com/aspose-cells/Aspose.Cells-for-Java
// Access first worksheet of gridweb
GridWorksheet sheet = gridweb.getWorkSheets().get(0);

// Access cell A1
GridCell cell = sheet.getCells().get("A1");

// Access hyperlink of cell A1 if it contains any
GridHyperlink lnk = sheet.getHyperlinks().getHyperlink(cell);

if (lnk == null) {
// This cell does not have any hyperlink
} else {
// This cell does have hyperlink, access its properties e.g. address
String addr = lnk.getAddress();
}