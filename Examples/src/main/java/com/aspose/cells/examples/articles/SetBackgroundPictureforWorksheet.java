package com.aspose.cells.examples.articles;

import java.io.FileInputStream;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;
import java.io.*;

public class SetBackgroundPictureforWorksheet {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SetBackgroundPictureforWorksheet.class) + "articles/";
		// Instantiate a new Workbook.
		Workbook workbook = new Workbook();
		// Get the first worksheet.
		Worksheet sheet = workbook.getWorksheets().get(0);

		// Get the image file.
		File file = new File(dataDir + "school.jpg");
		// Get the picture into the streams.
		byte[] imageData = new byte[(int) file.length()];
		FileInputStream fis = new FileInputStream(file);
		fis.read(imageData);

		// Set the background image for the sheet.
		sheet.setBackground(imageData);

		// Save the excel file
		workbook.save(dataDir + "SBPforWorksheet.xls");

	}
}
