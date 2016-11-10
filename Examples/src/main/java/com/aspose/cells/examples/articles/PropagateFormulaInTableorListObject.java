package com.aspose.cells.examples.articles;

import com.aspose.cells.ListObject;
import com.aspose.cells.TableStyleType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class PropagateFormulaInTableorListObject {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(PropagateFormulaInTableorListObject.class) + "articles/";
		// Create workbook object
		Workbook book = new Workbook();

		// Access first worksheet
		Worksheet sheet = book.getWorksheets().get(0);

		// Add column headings in cell A1 and B1
		sheet.getCells().get(0, 0).putValue("Column A");
		sheet.getCells().get(0, 1).putValue("Column B");

		// Add list object, set its name and style
		int idx = sheet.getListObjects().add(0, 0, 1, sheet.getCells().getMaxColumn(), true);
		ListObject listObject = sheet.getListObjects().get(idx);
		listObject.setTableStyleType(TableStyleType.TABLE_STYLE_MEDIUM_2);
		listObject.setDisplayName("Table");

		// Set the formula of second column so that it propagates to new rows
		// automatically while entering data
		listObject.getListColumns().get(1).setFormula("=[Column A] + 1");

		// Save the workbook in xlsx format
		book.save(dataDir + "PropagateFormulaInTable_out.xlsx");
	}
}
