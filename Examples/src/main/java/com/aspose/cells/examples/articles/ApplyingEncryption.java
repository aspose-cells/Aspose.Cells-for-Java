package com.aspose.cells.examples.articles;

import com.aspose.cells.EncryptionType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ApplyingEncryption {
	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(ApplyingEncryption.class) + "articles/";

		// Instantiate a Workbook object.
		Workbook workbook = new Workbook(dataDir + "Book1.xls");

		// Password protect the file.
		workbook.getSettings().setPassword("1234");

		// Specify Strong Encryption type (RC4,Microsoft Strong Cryptographic Provider).
		workbook.setEncryptionOptions(EncryptionType.STRONG_CRYPTOGRAPHIC_PROVIDER, 128);

		// Save the Excel file.
		workbook.save(dataDir + "AEncryption_out.xls");

	}
}
