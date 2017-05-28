package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class DisableCompatibilityChecker {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(DisableCompatibilityChecker.class) + "articles/";
		// Open a template file
		Workbook workbook = new Workbook(dataDir + "sample.xlsx");

		// Disable the compatibility checker
		workbook.getSettings().setCheckComptiliblity(false);

		// Saving the Excel file
		workbook.save(dataDir + "DCChecker_out.xls");

	}
}
