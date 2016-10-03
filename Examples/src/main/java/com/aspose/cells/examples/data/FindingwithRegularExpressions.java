package com.aspose.cells.examples.data;

import com.aspose.cells.Cells;
import com.aspose.cells.FindOptions;
import com.aspose.cells.LookAtType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class FindingwithRegularExpressions {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(FindingwithRegularExpressions.class) + "data/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Finding the cell containing the specified formula
		Cells cells = worksheet.getCells();

		// Instantiate FindOptions
		FindOptions findOptions = new FindOptions();

		// Instantiate FindOptions
		FindOptions opt = new FindOptions();
		// Set the search key of find() method as standard RegEx
		opt.setRegexKey(true);
		opt.setLookAtType(LookAtType.ENTIRE_CONTENT);
		cells.find("abc[\\s]*$", null, opt);
	}
}
