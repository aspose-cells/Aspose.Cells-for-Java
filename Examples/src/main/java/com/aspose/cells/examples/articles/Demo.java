package com.aspose.cells.examples.articles;

import com.aspose.cells.Cells;
import com.aspose.cells.OoxmlSaveOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;


public class Demo {
	private static final String OUTPUT_FILE_PATH = Utils.getSharedDataDir(Demo.class) + "articles/";

	public static void main(String[] args) throws Exception {
		// Instantiate a new Workbook
		Workbook wb = new Workbook();
		// set the sheet count
		int sheetCount = 1;
		// set the number of rows for the big matrix
		int rowCount = 100000;
		// specify the worksheet
		for (int k = 0; k < sheetCount; k++) {
			Worksheet sheet = null;
			if (k == 0) {
				sheet = wb.getWorksheets().get(k);
				sheet.setName("test");
			} else {
				int sheetIndex = wb.getWorksheets().add();
				sheet = wb.getWorksheets().get(sheetIndex);
				sheet.setName("test" + sheetIndex);
			}
			Cells cells = sheet.getCells();

			// set the columns width
			for (int j = 0; j < 15; j++) {
				cells.setColumnWidth(j, 15);
			}

			// traverse the columns for adding hyperlinks and merging
			for (int i = 0; i < rowCount; i++) {
				// The first 10 columns
				for (int j = 0; j < 10; j++) {
					if (j % 3 == 0) {
						cells.merge(i, j, 1, 2, false, false);
					}

					if (i % 50 == 0) {
						if (j == 0) {
							sheet.getHyperlinks().add(i, j, 1, 1, "test!A1");
						} else if (j == 3) {
							sheet.getHyperlinks().add(i, j, 1, 1, "http://www.google.com");
						}
					}
				}

				// The second 10 columns
				for (int j = 10; j < 20; j++) {
					if (j == 12) {
						cells.merge(i, j, 1, 3, false, false);
					}
				}
			}
		}

		// Create an object with respect to LightCells data provider
		LightCellsDataProviderDemo dataProvider = new LightCellsDataProviderDemo(wb, 1, rowCount, 20);
		// Specify the XLSX file's Save options
		OoxmlSaveOptions opt = new OoxmlSaveOptions();
		// Set the data provider for the file
		opt.setLightCellsDataProvider(dataProvider);

		// Save the big file
		wb.save(OUTPUT_FILE_PATH + "/Demo_out.xlsx", opt);
	}
}
