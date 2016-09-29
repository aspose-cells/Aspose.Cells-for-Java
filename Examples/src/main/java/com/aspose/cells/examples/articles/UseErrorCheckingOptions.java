package com.aspose.cells.examples.articles;

import com.aspose.cells.CellArea;
import com.aspose.cells.ErrorCheckOption;
import com.aspose.cells.ErrorCheckOptionCollection;
import com.aspose.cells.ErrorCheckType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class UseErrorCheckingOptions {
	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(UseErrorCheckingOptions.class) + "articles/";

		// Create a workbook and opening a template spreadsheet
		Workbook workbook = new Workbook(dataDir + "Book1.xls");

		// Get the first worksheet
		Worksheet sheet = workbook.getWorksheets().get(0);
		// Instantiate the error checking options
		ErrorCheckOptionCollection opts = sheet.getErrorCheckOptions();

		int index = opts.add();
		ErrorCheckOption opt = opts.get(index);
		// Disable the numbers stored as text option
		opt.setErrorCheck(ErrorCheckType.TEXT_NUMBER, false);
		// Set the range
		opt.addRange(CellArea.createCellArea(0, 0, 65535, 255));

		// Save the Excel file
		workbook.save(dataDir + "UseErrorCheckingOptions_out.xls");

	}
}
