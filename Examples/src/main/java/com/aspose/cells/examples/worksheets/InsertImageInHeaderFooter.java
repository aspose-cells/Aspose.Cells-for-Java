package com.aspose.cells.examples.worksheets;

import java.io.FileInputStream;

import com.aspose.cells.PageSetup;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class InsertImageInHeaderFooter {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(InsertImageInHeaderFooter.class) + "worksheets/";
		// Creating a Workbook object
		Workbook workbook = new Workbook();

		// Creating a string variable to store the url of the logo/picture
		String logo_url = dataDir + "school.jpg";

		// Creating the instance of the FileInputStream object to open the logo/picture in the stream
		FileInputStream inFile = new FileInputStream(logo_url);

		// Creating a PageSetup object to get the page settings of the first worksheet of the workbook
		PageSetup pageSetup = workbook.getWorksheets().get(0).getPageSetup();

		// Setting the logo/picture in the central section of the page header
		pageSetup.setHeader(1, "&G");
		byte[] picData = new byte[inFile.available()];
		inFile.read(picData);
		pageSetup.setHeaderPicture(1, picData);

		// Setting the Sheet's name in the right section of the page header with the script
		pageSetup.setHeader(2, "&A");

		// Saving the workbook
		workbook.save(dataDir + "InsertImageInHeaderFooter_out.xls");

		// Closing the FileStream object
		inFile.close();
	}
}
