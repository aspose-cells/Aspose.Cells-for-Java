package com.aspose.cells.examples.data;

import java.util.Iterator;

import com.aspose.cells.BorderType;
import com.aspose.cells.Cell;
import com.aspose.cells.CellBorderType;
import com.aspose.cells.Color;
import com.aspose.cells.Range;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ConvertCellsAddresstoRangeorCellArea {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ConvertCellsAddresstoRangeorCellArea.class) + "data/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Obtaining the reference of the newly added worksheet
		int sheetIndex = workbook.getWorksheets().add();
		Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);

		// Accessing the "A1" cell from the worksheet
		Cell cell = worksheet.getCells().get("A1");

		// Adding some value to the "A1" cell
		cell.setValue("Hello World!");

		// Creating a range of cells based on cells Address.
		Range range = worksheet.getCells().createRange("A1:F10");

		// Specify a Style object for borders.
		Style style = cell.getStyle();
		// Setting the line style of the top border

		style.setBorder(BorderType.TOP_BORDER, CellBorderType.THICK, Color.getBlack());

		style.setBorder(BorderType.BOTTOM_BORDER, CellBorderType.THICK, Color.getBlack());

		style.setBorder(BorderType.LEFT_BORDER, CellBorderType.THICK, Color.getBlack());

		style.setBorder(BorderType.RIGHT_BORDER, CellBorderType.THICK, Color.getBlack());

		Iterator cellArray = range.iterator();
		while (cellArray.hasNext()) {
			Cell temp = (Cell) cellArray.next();
			// Saving the modified style to the cell.
			temp.setStyle(style);
		}

		// Saving the Excel file
		workbook.save(dataDir + "CCAToROrCArea_out.xls");

	}
}
