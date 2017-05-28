package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class ExceltoPDF {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ExceltoPDF.class) + "TechnicalArticles/";
		// Initialize a new Workbook
		// Open an Excel file
		Workbook workbook = new Workbook(dataDir + "Mybook.xls");

		// Implement one page per worksheet option
		PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
		pdfSaveOptions.setOnePagePerSheet(true);

		// Save the PDF file
		workbook.save(dataDir + "ExceltoPDF_out.pdf", pdfSaveOptions);

	}
}
