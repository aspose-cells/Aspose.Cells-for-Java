package com.aspose.cells.examples.data;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Font;
import com.aspose.cells.FontUnderlineType;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SettingFontUnderlineType {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SettingFontUnderlineType.class) + "data/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Accessing the added worksheet in the Excel file
		int sheetIndex = workbook.getWorksheets().add();
		Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);
		Cells cells = worksheet.getCells();

		// Adding some value to the "A1" cell
		Cell cell = cells.get("A1");
		cell.setValue("Hello Aspose!");

		// Setting the font to be underlined
		Style style = cell.getStyle();
		Font font = style.getFont();
		font.setUnderline(FontUnderlineType.SINGLE);
		cell.setStyle(style);

		cell.setStyle(style);

		// Saving the modified Excel file in default format
		workbook.save(dataDir + "SFUnderlineType_out.xls");
	}
}
