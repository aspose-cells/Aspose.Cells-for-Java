package com.aspose.cells.examples.data;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class AddingLinkToAnotherCell {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddingLinkToAnotherCell.class) + "data/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Obtaining the reference of the first worksheet.
		WorksheetCollection worksheets = workbook.getWorksheets();
		workbook.getWorksheets().add();
		Worksheet sheet = worksheets.get(0);

		// Setting a value to the "A1" cell
		Cells cells = sheet.getCells();
		Cell cell = cells.get("A1");
		cell.setValue("Visit Aspose");

		// Setting the font color of the cell to Blue
		Style style = cell.getStyle();
		style.getFont().setColor(Color.getBlue());

		// Setting the font of the cell to Single Underline
		style.getFont().setUnderline(FontUnderlineType.SINGLE);
		cell.setStyle(style);

		HyperlinkCollection hyperlinks = sheet.getHyperlinks();

		// Adding an internal hyperlink to the "B9" cell of the other worksheet "Sheet2" in the same Excel file
		hyperlinks.add("B3", 1, 1, "Sheet2!B9");

		// Saving the Excel file
		workbook.save(dataDir + "ALinkTACell_out.xls");

		// Print message
		System.out.println("Process completed successfully");

	}
}
