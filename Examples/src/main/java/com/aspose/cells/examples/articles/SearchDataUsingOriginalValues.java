package com.aspose.cells.examples.articles;

import com.aspose.cells.Cell;
import com.aspose.cells.FindOptions;
import com.aspose.cells.LookAtType;
import com.aspose.cells.LookInType;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SearchDataUsingOriginalValues {
	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(SearchDataUsingOriginalValues.class) + "articles/";
		// Create workbook object
		Workbook workbook = new Workbook();

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Add 10 in cell A1 and A2
		worksheet.getCells().get("A1").putValue(10);
		worksheet.getCells().get("A2").putValue(10);

		// Add Sum formula in cell D4 but customize it as ---
		Cell cell = worksheet.getCells().get("D4");

		Style style = cell.getStyle();
		style.setCustom("---");
		cell.setStyle(style);

		// The result of formula will be 20, but 20 will not be visible because the cell is formated as ---
		cell.setFormula("=Sum(A1:A2)");

		// Calculate the workbook
		workbook.calculateFormula();

		// Create find options, we will search 20 using. original values otherwise 20 will never be found,because it is formatted
		// as
		FindOptions options = new FindOptions();
		options.setLookInType(LookInType.ORIGINAL_VALUES);
		options.setLookAtType(LookAtType.ENTIRE_CONTENT);

		Cell foundCell = null;
		Object obj = 20;

		// Find 20 which is Sum(A1:A2) and formatted as ---
		foundCell = worksheet.getCells().find(obj, foundCell, options);

		// Print the found cell
		System.out.println(foundCell);

		// Save the workbook
		workbook.save(dataDir + "SDUOriginalValues_out.xlsx");

	}
}
