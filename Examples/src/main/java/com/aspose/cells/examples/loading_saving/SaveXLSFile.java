package com.aspose.cells.examples.loading_saving;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class SaveXLSFile {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SaveXLSFile.class) + "loading_saving/";

		// Creating an Workbook object with an Excel file path
		Workbook workbook = new Workbook();

		// Save in xls format
		workbook.save(dataDir + "SXLSFile_out.xls", FileFormatType.EXCEL_97_TO_2003);

		// Print Message
		System.out.println("Worksheets are saved successfully.");

	}
}
