package com.aspose.cells.examples.articles;

import java.io.File;

import com.aspose.cells.CellsHelper;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

//ExStart:TestIFilePathProvider

public class TestIFilePathProvider {
	// This is the directory path which contains the sample.xlsx file
	static String dataDir = Utils.getSharedDataDir(TestIFilePathProvider.class) + "articles/";

	public static void main(String[] args) throws Exception {

		/*
		 * If you will not set the license, program will go in infinite loop because Aspose.Cells will always make the warning
		 * worksheet as active sheet in Evaluation mode.
		 */
		SetLicense();

		// Check if license is set, otherwise do not proceed
		Workbook wb = new Workbook();

		if (wb.isLicensed() == false) {
			System.out.println("You must set the license to execute this code successfully.");
		} else {
			// Test IFilePathProvider interface
			TestIFilePathProvider pg = new TestIFilePathProvider();
			pg.TestFilePathProvider();
			System.out.println("Done.");
		}
	}

	static void SetLicense() throws Exception {
		String licPath = dataDir + "Aspose.Total.Java.lic";
		com.aspose.cells.License lic = new com.aspose.cells.License();
		lic.setLicense(licPath);

		System.out.println(CellsHelper.getVersion());
	}

	void TestFilePathProvider() throws Exception {
		// Create subdirectory for second and third worksheets
		File dir = new File(dataDir + "OtherSheets");
		dir.mkdir();

		// Load sample workbook from your directory
		Workbook wb = new Workbook(dataDir + "Sample.xlsx");

		// Save worksheets to separate html files
		// Because of IFilePathProvider, hyperlinks will not be broken.
		for (int i = 0; i < wb.getWorksheets().getCount(); i++) {
			// Set the active worksheet to current value of variable i
			wb.getWorksheets().setActiveSheetIndex(i);

			// Creat html save option
			ImplementingIStreamProvider options = new ImplementingIStreamProvider();
			options.setExportActiveWorksheetOnly(true);

			// If you will comment this line, then hyperlinks will be broken
			options.setFilePathProvider(new FilePathProvider());

			// Sheet actual index which starts from 1 not from 0
			int sheetIndex = i + 1;

			String filePath = "";

			// Save first sheet to same directory and second and third
			// worksheets to subdirectory
			if (i == 0) {
				filePath = dataDir + "Sheet1.html";
			} else {
				filePath = dataDir + "Sheet" + sheetIndex + ".html";
			}

			// Save the worksheet to html file
			wb.save(filePath);
		}
	}
}

