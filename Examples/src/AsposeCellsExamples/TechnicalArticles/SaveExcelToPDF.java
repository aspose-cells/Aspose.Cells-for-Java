package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.PdfOptimizationType;
import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class SaveExcelToPDF {
	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(SaveExcelToPDF.class) + "TechnicalArticles/";
		// Load excel file into workbook object
		Workbook workbook = new Workbook(dataDir + "sample.xlsx");
		// Save into Pdf with Minimum size
		PdfSaveOptions opts = new PdfSaveOptions();
		opts.setOptimizationType(PdfOptimizationType.MINIMUM_SIZE);
		workbook.save(dataDir + "SExcelToPDF_out.pdf", opts);

	}
}
