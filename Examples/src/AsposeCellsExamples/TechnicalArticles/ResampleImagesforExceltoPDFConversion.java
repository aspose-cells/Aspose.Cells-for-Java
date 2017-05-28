package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class ResampleImagesforExceltoPDFConversion {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ResampleImagesforExceltoPDFConversion.class) + "TechnicalArticles/";
		// Initialize a new Workbook
		// Open an Excel file
		Workbook workbook = new Workbook(dataDir + "Book1.xlsx");

		// Instantiate the PdfSaveOptions
		PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
		// Set Image Resample properties
		pdfSaveOptions.setImageResample(300, 70);

		// Save the PDF file
		workbook.save(dataDir + "ReSIfEToPDFC_out.pdf", pdfSaveOptions);

	}
}
