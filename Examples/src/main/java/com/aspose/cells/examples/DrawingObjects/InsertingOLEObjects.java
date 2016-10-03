package com.aspose.cells.examples.DrawingObjects;

import java.io.FileInputStream;

import com.aspose.cells.OleObject;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;
import java.io.*;

public class InsertingOLEObjects {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(InsertingOLEObjects.class) + "DrawingObjects/";

		// Get the image file.
		File file = new File(dataDir + "logo.jpg");

		// Get the picture into the streams.
		byte[] img = new byte[(int) file.length()];
		FileInputStream fis = new FileInputStream(file);
		fis.read(img);

		// Get the excel file into the streams.
		file = new File(dataDir + "Book1.xls");
		byte[] data = new byte[(int) file.length()];
		fis = new FileInputStream(file);
		fis.read(data);

		// Instantiate a new Workbook.
		Workbook wb = new Workbook();

		// Get the first worksheet.
		Worksheet sheet = wb.getWorksheets().get(0);

		// Add an Ole object into the worksheet with the image shown in MS Excel.
		int oleObjIndex = sheet.getOleObjects().add(14, 3, 200, 220, img);
		OleObject oleObj = sheet.getOleObjects().get(oleObjIndex);

		// Set embedded ole object data.
		oleObj.setObjectData(data);

		// Save the excel file
		wb.save(dataDir + "InsertingOLEObjects_out.xls");
	}
}
