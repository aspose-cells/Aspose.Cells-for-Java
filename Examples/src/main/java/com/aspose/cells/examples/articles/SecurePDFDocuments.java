package com.aspose.cells.examples.articles;

import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.PdfSecurityOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class SecurePDFDocuments {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SecurePDFDocuments.class) + "articles/";
		// Open an Excel file
		Workbook workbook = new Workbook(dataDir + "input.xlsx");

		// Instantiate PDFSaveOptions to manage security attributes
		PdfSaveOptions saveOption = new PdfSaveOptions();

		saveOption.setSecurityOptions(new PdfSecurityOptions());
		// Set the user password
		saveOption.getSecurityOptions().setUserPassword("user");

		// Set the owner password
		saveOption.getSecurityOptions().setOwnerPassword("owner");

		// Disable extracting content permission
		saveOption.getSecurityOptions().setExtractContentPermission(false);

		// Disable print permission
		saveOption.getSecurityOptions().setPrintPermission(false);

		// Save the PDF document with encrypted settings
		workbook.save(dataDir + "SecurePDFDocuments_out.pdf", saveOption);

	}
}
