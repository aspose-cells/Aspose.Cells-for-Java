package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

import java.io.FileOutputStream;

public class ExportCustomPropertiesToPDF {

	// The path to the documents directory.
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		// ExStart:ExportCustomPropertiesToPDF
		// Load template workbook 
		Workbook workbook = new Workbook(srcDir + "sourceWithCustProps.xlsx");

		// Create an instance of PdfSaveOptions and pass SaveFormat to the constructor
		PdfSaveOptions pdfSaveOpt = new PdfSaveOptions(SaveFormat.PDF);

		// Set CustomPropertiesExport property to PdfCustomPropertiesExport.Standard
		pdfSaveOpt.setCustomPropertiesExport(PdfCustomPropertiesExport.STANDARD);

		// Save the workbook to PDF format while passing the object of PdfSaveOptions
		workbook.save(outDir + "outSourceWithCustProps.pdf", pdfSaveOpt);
		
		// Print message
		System.out.println("Export Custom Properties To PDF performed successfully.");
		// ExEnd:ExportCustomPropertiesToPDF		
	}
}
