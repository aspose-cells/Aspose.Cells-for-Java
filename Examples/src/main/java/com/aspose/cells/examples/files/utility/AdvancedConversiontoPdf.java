package com.aspose.cells.examples.files.utility;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

import java.io.FileOutputStream;

public class AdvancedConversiontoPdf {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AdvancedConversiontoPdf.class) + "files/";

		// Create a new Workbook.
		Workbook workbook = new Workbook();
		Cells cell = workbook.getWorksheets().get(0).getCells();
		cell.get("A12").setValue("Test PDF");
		PdfSaveOptions pdfOptions = new PdfSaveOptions();

		pdfOptions.setCompliance(PdfCompliance.PDF_A_1_B);
		workbook.save(dataDir + "ACToPdf-out.pdf", pdfOptions);

		// Print message
		System.out.println("Advanced Conversion performed successfully.");


	}
}
