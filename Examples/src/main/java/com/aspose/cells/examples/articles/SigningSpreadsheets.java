package com.aspose.cells.examples.articles;

import java.io.FileInputStream;
import java.io.InputStream;
import java.security.KeyStore;

import com.aspose.cells.DateTime;
import com.aspose.cells.DigitalSignature;
import com.aspose.cells.DigitalSignatureCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class SigningSpreadsheets {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SigningSpreadsheets.class) + "articles/";
		// Create an instance of DigitalSignatureCollection
		DigitalSignatureCollection signatures = new DigitalSignatureCollection();

		// Load the certificate into an instance of InputStream
		InputStream inStream = new FileInputStream("d:/temp.pfx");

		// Create an instance of KeyStore with PKCS12 cryptography
		KeyStore inputKeyStore = KeyStore.getInstance("PKCS12");

		// Use the KeyStore.load method to load the certificate stream and its password
		inputKeyStore.load(inStream, KEYSTORE_PASSWORD.toCharArray());

		// Create an instance of DigitalSignature and pass the instance of KeyStore, password, comments and time
		DigitalSignature signature = new DigitalSignature(inputKeyStore, KEYSTORE_PASSWORD, "test for  sign",
				DateTime.getNow());

		// Add the instance of DigitalSignature into the collection
		signatures.add(signature);

		// Load an existing spreadsheet using the Workbook class
		Workbook workbook = new Workbook(dataDir + "unsigned.xlsx");

		// Set the signature
		workbook.setDigitalSignature(signatures);

		// Save the signed spreadsheet
		workbook.save(dataDir + "SSpreadsheets_out.xlsx");

	}
}
