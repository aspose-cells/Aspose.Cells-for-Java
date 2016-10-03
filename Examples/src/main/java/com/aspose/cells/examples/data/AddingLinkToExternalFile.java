package com.aspose.cells.examples.data;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.FontUnderlineType;
import com.aspose.cells.HyperlinkCollection;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class AddingLinkToExternalFile {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddingLinkToExternalFile.class) + "data/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Obtaining the reference of the first worksheet.
		WorksheetCollection worksheets = workbook.getWorksheets();
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

		// Adding a link to the external file
		hyperlinks.add("A5", 1, 1, dataDir + "book1.xls");

		// Saving the Excel file
		workbook.save(dataDir + "ALToEFile_out.xls");

		// Print message
		System.out.println("Process completed successfully");

	}
}
