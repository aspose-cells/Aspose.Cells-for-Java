package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class CalculationofExcelMINIFSandMAXIFSfunctions {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CalculationofExcelMINIFSandMAXIFSfunctions.class) + "articles/";

		// Load your source workbook containing MINIFS and MAXIFS functions
		Workbook wb = new Workbook(dataDir + "sample_MINIFS_MAXIFS.xlsx");

		// Perform Aspose.Cells formula calculation
		wb.calculateFormula();

		// Save the calculations result in pdf format
		PdfSaveOptions opts = new PdfSaveOptions();
		opts.setOnePagePerSheet(true);
		wb.save(dataDir + "CalculationofExcel_out.pdf", opts);
	}
}
