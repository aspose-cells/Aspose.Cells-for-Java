package com.aspose.cells.examples.worksheets;

import java.io.FileInputStream;

import com.aspose.cells.PageSetup;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class SetHeadersAndFooters {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(SetHeadersAndFooters.class) + "worksheets/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Obtaining the reference of the PageSetup of the worksheet
		PageSetup pageSetup = workbook.getWorksheets().get(0).getPageSetup();

		// Setting worksheet name at the left header
		pageSetup.setHeader(0, "&A");

		// Setting current date and current time at the central header and changing the font of the header
		pageSetup.setHeader(1, "&\"Times New Roman,Bold\"&D-&T");

		// Setting current file name at the right header and changing the font of the header
		pageSetup.setHeader(2, "&\"Times New Roman,Bold\"&12&F");

		// Setting a string at the left footer and changing the font of the footer
		pageSetup.setFooter(0, "Hello World! &\"Courier New\"&14 123");

		// Setting picture at the central footer
		pageSetup.setFooter(1, "&G");

		FileInputStream fis = new FileInputStream(dataDir + "footer.jpg");
		byte[] picData = new byte[fis.available()];
		fis.read(picData);
		pageSetup.setFooterPicture(1, picData);
		fis.close();

		// Setting the current page number and page count at the right footer
		pageSetup.setFooter(2, "&Pof&N");
		workbook.save(dataDir + "SetHeadersAndFooters_out.xls");
	}
}
