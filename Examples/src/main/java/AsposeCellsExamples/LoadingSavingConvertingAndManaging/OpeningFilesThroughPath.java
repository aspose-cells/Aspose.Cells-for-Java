package com.aspose.cells.examples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

import java.io.FileInputStream;

public class OpeningFilesThroughPath {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(OpeningFilesThroughPath.class) + "LoadingSavingConvertingAndManaging/";

		// Opening from path.
		// Creating an Workbook object with an Excel file path
		Workbook workbook1 = new Workbook(dataDir + "Book1.xlsx");

		// Print message
		System.out.println("Workbook opened using path successfully.");

	}
}
