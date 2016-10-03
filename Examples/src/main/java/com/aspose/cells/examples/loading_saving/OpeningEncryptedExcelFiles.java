package com.aspose.cells.examples.loading_saving;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

import java.io.FileInputStream;

public class OpeningEncryptedExcelFiles {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(OpeningEncryptedExcelFiles.class) + "loading_saving/";

		// Opening Encrypted Excel Files
		// Creating and EXCEL_97_TO_2003 LoadOptions object
		LoadOptions loadOptions6 = new LoadOptions(FileFormatType.EXCEL_97_TO_2003);

		// Setting the password for the encrypted Excel file
		loadOptions6.setPassword("1234");

		// Creating an Workbook object with file path and the loadOptions object
		Workbook workbook8 = new Workbook(dataDir + "encryptedBook.xls", loadOptions6);

		// Print message
		System.out.println("Encrypted workbook has been opened successfully.");

	}
}
