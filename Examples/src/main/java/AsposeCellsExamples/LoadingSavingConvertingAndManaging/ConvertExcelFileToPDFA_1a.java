package com.aspose.cells.examples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class ConvertExcelFileToPDFA_1a {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ConvertExcelFileToPDFA_1a.class) + "LoadingSavingConvertingAndManaging/";
	
		//Create workbook object
		Workbook wb = new Workbook();
		 
		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);
		 
		//Access cell B5 and add some message inside it
		Cell cell = ws.getCells().get("B5");
		cell.putValue("This PDF format is compatible with PDFA-1a.");
		 
		//Create pdf save options and set its compliance to PDFA-1a
		PdfSaveOptions opts = new PdfSaveOptions();
		opts.setCompliance(PdfCompliance.PDF_A_1_A);
		 
		//Save the output pdf
		wb.save(dataDir + "outputCompliancePdfA1a.pdf", opts);

		// Print message
		System.out.println("ConvertExcelFileToPDFA_1a executed successfully.");

	}
}
