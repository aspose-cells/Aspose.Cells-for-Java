package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

import java.io.FileOutputStream;

public class ContentCopyForAccessibilityOption {
	// The path to the documents directory.
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();


	public static void main(String[] args) throws Exception {

		// ExStart:ContentCopyForAccessibilityOption
		// Create a new Workbook.
		Workbook workbook = new Workbook(srcDir + "sourseSampleCountryNames.xlsx");
		Cells cell = workbook.getWorksheets().get(0).getCells();
		cell.get("A12").setValue("Test PDF");
		PdfSaveOptions pdfOptions = new PdfSaveOptions();

		pdfOptions.setCompliance(PdfCompliance.PDF_A_1_B);
		workbook.save(outDir + "ACToPdf_out.pdf", pdfOptions);

		// Print message
		System.out.println("Content Copy For Accessibility Option performed successfully.");
		// ExEnd:ContentCopyForAccessibilityOption	

	}
}
