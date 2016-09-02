package com.aspose.cells.examples.articles;

import com.aspose.cells.DigitalSignature;
import com.aspose.cells.DigitalSignatureCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ValidatingSpreadsheets {
	public static void main(String[] args) throws Exception {
		// ExStart:ValidatingSpreadsheets
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(ValidatingSpreadsheets.class);
		// Load an existing spreadsheet in an instance of Workbook
		Workbook workbook = new Workbook(dataDir + "signed.xlsx");

		// Retrieve the collection of digital signatures from the Workbook
		DigitalSignatureCollection signatures = workbook.getDigitalSignature();

		// Loop over the collection of digital signatures
		for (DigitalSignature signature : (Iterable<DigitalSignature>) signatures) {
			// Check the signature status using the isValid property
			System.out.println(signature.isValid());
		}
		// ExEnd:ValidatingSpreadsheets
	}
}
