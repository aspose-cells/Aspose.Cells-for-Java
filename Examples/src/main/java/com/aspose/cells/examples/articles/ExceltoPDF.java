package com.aspose.cells.examples.articles;

import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ExceltoPDF {
	public static void main(String[] args) throws Exception {
		// ExStart:ExceltoPDF
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(ExceltoPDF.class);
		// Initialize a new Workbook
		// Open an Excel file
		Workbook workbook = new Workbook(dataDir + "Mybook.xls");

		// Implement one page per worksheet option
		PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
		pdfSaveOptions.setOnePagePerSheet(true);

		// Save the PDF file
		workbook.save(dataDir + "OutputFile.pdf", pdfSaveOptions);
		// ExEnd:ExceltoPDF
	}
}
