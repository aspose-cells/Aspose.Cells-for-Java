package com.aspose.cells.examples.worksheets;

import java.io.FileInputStream;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.PageSetup;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class InsertGraphicinFirstPageHeaderOnly {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(InsertGraphicinFirstPageHeaderOnly.class) + "worksheets/";

		// Creating a Workbook object
		Workbook workbook = new Workbook();

		// Get the first worksheet (default).
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Adding some sample value to cells
		Cells cells = worksheet.getCells();
		Cell cell = cells.get("A1");
		cell.setValue("Page1");
		cell = cells.get("A60");
		cell.setValue("Page2");
		cell = cells.get("A113");
		cell.setValue("Page3");

		// Creating a PageSetup object to get the page settings of the first
		// worksheet of the workbook
		PageSetup pageSetup = worksheet.getPageSetup();

		// Creating a string variable to store the url of the logo/picture
		String logo_url = dataDir + "school.jpg";

		// Creating the instance of the FileInputStream object to open the logo/picture in the stream
		FileInputStream inFile = new FileInputStream(logo_url);
		byte[] picData = new byte[inFile.available()];
		inFile.read(picData);

		// Setting the logo/picture in the right section of the first page header only
		pageSetup.setHFDiffFirst(true);
		pageSetup.setFirstPageHeader(2, "&G");
		pageSetup.setPicture(true, false, true, 2, picData);

		// Saving the workbook
		workbook.save(dataDir + "IGInFirstPageHeaderOnly_out.xlsx");

		// Closing the FileStream object
		inFile.close();
	}
}
