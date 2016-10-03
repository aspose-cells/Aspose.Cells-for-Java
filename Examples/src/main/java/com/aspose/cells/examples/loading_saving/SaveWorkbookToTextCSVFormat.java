package com.aspose.cells.examples.loading_saving;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;

public class SaveWorkbookToTextCSVFormat {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SaveWorkbookToTextCSVFormat.class) + "loading_saving/";

		// Load your source workbook
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// 0-byte array
		byte[] workbookData = new byte[0];

		// Text save options. You can use any type of separator
		TxtSaveOptions opts = new TxtSaveOptions();
		opts.setSeparator('\t');

		// Copy each worksheet data in text format inside workbook data array
		for (int idx = 0; idx < workbook.getWorksheets().getCount(); idx++) {
			// Save the active worksheet into text format
			ByteArrayOutputStream bout = new ByteArrayOutputStream();
			workbook.getWorksheets().setActiveSheetIndex(idx);
			workbook.save(bout, opts);

			// Save the worksheet data into sheet data array
			byte[] sheetData = bout.toByteArray();

			// Combine this worksheet data into workbook data array
			byte[] combinedArray = new byte[workbookData.length + sheetData.length];
			System.arraycopy(workbookData, 0, combinedArray, 0, workbookData.length);
			System.arraycopy(sheetData, 0, combinedArray, workbookData.length, sheetData.length);

			workbookData = combinedArray;
		}

		// Save entire workbook data into file
		FileOutputStream fout = new FileOutputStream(dataDir + "SWTTextCSVFormat-out.txt");
		fout.write(workbookData);
		fout.close();

		// Print message
		System.out.println("Excel to Text File Conversion performed successfully.");

	}
}
