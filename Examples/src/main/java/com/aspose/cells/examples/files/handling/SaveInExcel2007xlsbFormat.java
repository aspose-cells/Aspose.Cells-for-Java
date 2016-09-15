package com.aspose.cells.examples.files.handling;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class SaveInExcel2007xlsbFormat {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SaveInExcel2007xlsbFormat.class) + "files/";

		// Creating an Workbook object with an Excel file path
		Workbook workbook = new Workbook();

		// Save in Excel2007 xlsb format
		workbook.save(dataDir + "SIE2007xlsbFormat-out.xlsb", FileFormatType.XLSB);

		// Print Message
		System.out.println("Worksheets are saved successfully.");

	}
}
