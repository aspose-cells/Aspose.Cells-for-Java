package com.aspose.cells.examples.files.utility;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class RenderUnicodeSupplimentaryCharacterToPDF {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(RenderUnicodeSupplimentaryCharacterToPDF.class) + "files/";
		// Load your source excel file containing Unicode Supplementary
		// characters
		Workbook wb = new Workbook(dataDir + "unicode-supplementary-characters.xlsx");

		// Save the workbook
		wb.save(dataDir + "RUSCToPDF-out.pdf");

	}

}
