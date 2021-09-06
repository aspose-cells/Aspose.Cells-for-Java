package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.LoadFormat;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.Workbook;

import AsposeCellsExamples.Utils;

public class OpeningSpreadsheetMLFiles {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(OpeningSpreadsheetMLFiles.class) + "LoadingSavingConvertingAndManaging/";

		// Opening SpreadsheetML Files
		// Creating and EXCEL_2003_XML LoadOptions object
		LoadOptions loadOptions3 = new LoadOptions(LoadFormat.SPREADSHEET_ML);

		// Creating an Workbook object with SpreadsheetML file path and the
		// loadOptions object
		new Workbook(dataDir + "Book3.xml", loadOptions3);

		// Print message
		System.out.println("SpreadSheetML format workbook has been opened successfully.");


	}
}
