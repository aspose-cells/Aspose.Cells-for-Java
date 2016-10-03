package com.aspose.cells.examples.loading_saving;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.TxtSaveOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.XpsSaveOptions;
import com.aspose.cells.examples.Utils;

import java.io.FileOutputStream;

public class SavingTextFilewithCustomSeparator {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SavingTextFilewithCustomSeparator.class) + "loading_saving/";

		// Creating an Workbook object with an Excel file path
		Workbook workbook = new Workbook(dataDir + "Book1.xlsx");

		TxtSaveOptions toptions = new TxtSaveOptions();
		// Specify the separator
		toptions.setSeparator(';');
		workbook.save(dataDir + "STFWCSeparator_out.csv");

		// Print Message
		System.out.println("Worksheets are saved successfully.");

	}
}
