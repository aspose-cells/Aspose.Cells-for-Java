package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class FitAllWorksheetColumns {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(FitAllWorksheetColumns.class) + "TechnicalArticles/";
		// Create and initialize an instance of Workbook
		Workbook book = new Workbook(dataDir + "TestBook.xlsx");
		// Create and initialize an instance of PdfSaveOptions
		PdfSaveOptions saveOptions = new PdfSaveOptions();
		// Set AllColumnsInOnePagePerSheet to true
		saveOptions.setAllColumnsInOnePagePerSheet(true);
		// Save Workbook to PDF fromart by passing the object of PdfSaveOptions
		book.save(dataDir + "FAWorksheetColumns_out.pdf", saveOptions);

	}
}
