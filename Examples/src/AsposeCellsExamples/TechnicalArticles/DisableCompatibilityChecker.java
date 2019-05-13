package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class DisableCompatibilityChecker {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(DisableCompatibilityChecker.class) + "TechnicalArticles/";
		// Open a template file
		Workbook workbook = new Workbook(dataDir + "book1.xlsx");

		// Disable the compatibility checker
		workbook.getSettings().setCheckCompatibility(false);

		// Saving the Excel file
		workbook.save(dataDir + "DCChecker_out.xls");
        // ExEnd:1

		System.out.println("DisableCompatibilityChecker executed successfully.");
	}
}
