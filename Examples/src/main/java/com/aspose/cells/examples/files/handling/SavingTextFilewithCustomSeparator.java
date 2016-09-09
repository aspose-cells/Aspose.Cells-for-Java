package com.aspose.cells.examples.files.handling;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.TxtSaveOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.XpsSaveOptions;
import com.aspose.cells.examples.Utils;

import java.io.FileOutputStream;

public class SavingTextFilewithCustomSeparator {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SavingTextFilewithCustomSeparator.class) + "files/";

		// Creating an Workbook object with an Excel file path
		Workbook workbook = new Workbook(dataDir + "Book1.xlsx");

		TxtSaveOptions toptions = new TxtSaveOptions();
		// Specify the separator
		toptions.setSeparator(';');
		workbook.save(dataDir + "STFWCSeparator-out.csv");

		// Print Message
		System.out.println("Worksheets are saved successfully.");
		// ExEnd:1
	}
}
