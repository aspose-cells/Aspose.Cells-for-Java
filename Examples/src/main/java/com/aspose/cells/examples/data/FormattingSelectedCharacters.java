package com.aspose.cells.examples.data;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.Font;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class FormattingSelectedCharacters {
	public static void main(String[] args) throws Exception {
		// Path to source file
		String dataDir = Utils.getSharedDataDir(FormattingSelectedCharacters.class) + "data/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Accessing the added worksheet in the Excel file
		int sheetIndex = workbook.getWorksheets().add();
		Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);
		Cells cells = worksheet.getCells();

		// Adding some value to the "A1" cell
		Cell cell = cells.get("A1");
		cell.setValue("Visit Aspose!");

		Font font = cell.characters(6, 7).getFont();

		// Setting the font of selected characters to bold
		font.setBold(true);

		// Setting the font color of selected characters to blue
		font.setColor(Color.getBlue());

		// Saving the Excel file
		workbook.save(dataDir + "FSCharacters_out.xls");
	}
}
