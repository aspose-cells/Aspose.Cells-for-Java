package com.aspose.cells.examples.files.handling;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;
import java.io.FileInputStream;

public class OpeningFiles {

	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(OpeningEncryptedExcelFiles.class) + "files/";
		// ExSart:1

		// 1. Opening from path.Creating an Workbook object with an Excel file path

		Workbook workbook1 = new Workbook(dataDir + "Book1.xls");

		// Print message
		System.out.println("Workbook opened using path successfully.");

		/*
		 * 2. Opening workbook from stream. Create a Stream object
		 */
		FileInputStream fstream = new FileInputStream(dataDir + "Book2.xls");

		// Creating an Workbook object with the stream object
		Workbook workbook2 = new Workbook(fstream);

		fstream.close();

		// Print message
		System.out.println("Workbook opened using stream successfully.");

		/*
		 * 3. Opening Microsoft Excel 97 Files Createing and EXCEL_97_TO_2003
		 * LoadOptions object
		 */
		LoadOptions loadOptions1 = new LoadOptions(FileFormatType.EXCEL_97_TO_2003);

		// Creating an Workbook object with excel 97 file path and the
		// loadOptions object
		Workbook workbook3 = new Workbook(dataDir + "Book_Excel97_2003.xls", loadOptions1);

		// Print message
		System.out.println("Excel 97 Workbook opened successfully.");

		/*
		 * 4. Opening Microsoft Excel 2007 XLSX Files Createing and XLSX
		 * LoadOptions object
		 */
		LoadOptions loadOptions2 = new LoadOptions(FileFormatType.XLSX);

		// Creating an Workbook object with 2007 xlsx file path and the
		// loadOptions object
		Workbook workbook4 = new Workbook(dataDir + "Book_Excel2007.xlsx", loadOptions2);

		// Print message
		System.out.println("Excel 2007 Workbook opened successfully.");

		/*
		 * 5. Opening SpreadsheetML Files Creating and EXCEL_2003_XML
		 * LoadOptions object
		 */
		LoadOptions loadOptions3 = new LoadOptions(FileFormatType.EXCEL_2003_XML);

		// Creating an Workbook object with SpreadsheetML file path and the
		// loadOptions object
		Workbook workbook5 = new Workbook(dataDir + "Book3.xml", loadOptions3);

		// Print message
		System.out.println("SpreadSheetML format workbook has been opened successfully.");

		/*
		 * 6. Opening CSV Files Creating and CSV LoadOptions object
		 */
		LoadOptions loadOptions4 = new LoadOptions(FileFormatType.CSV);

		// Creating an Workbook object with CSV file path and the loadOptions
		// object
		Workbook workbook6 = new Workbook(dataDir + "Book_CSV.csv", loadOptions4);

		// Print message
		System.out.println("CSV format workbook has been opened successfully.");

		/*
		 * 7. Opening Tab Delimited Files Creating and TAB_DELIMITED LoadOptions
		 * object
		 */
		LoadOptions loadOptions5 = new LoadOptions(FileFormatType.TAB_DELIMITED);

		// Creating an Workbook object with Tab Delimited text file path and the
		// loadOptions object
		Workbook workbook7 = new Workbook(dataDir + "Book1TabDelimited.txt", loadOptions5);

		// Print message
		System.out.println("Tab Delimited workbook has been opened successfully.");

		/*
		 * 8. Opening Encrypted Excel Files Creating and EXCEL_97_TO_2003
		 * LoadOptions object
		 */
		LoadOptions loadOptions6 = new LoadOptions(FileFormatType.EXCEL_97_TO_2003);

		// Setting the password for the encrypted Excel file
		loadOptions6.setPassword("1234");

		// Creating an Workbook object with file path and the loadOptions object
		Workbook workbook8 = new Workbook(dataDir + "encryptedBook.xls", loadOptions6);

		// Print message
		System.out.println("Encrypted workbook has been opened successfully.");
		// ExEnd:1
	}
}
