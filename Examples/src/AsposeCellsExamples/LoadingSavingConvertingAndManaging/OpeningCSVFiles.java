package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class OpeningCSVFiles {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(OpeningCSVFiles.class) + "LoadingSavingConvertingAndManaging/";
		// Opening CSV Files
		// Creating and CSV LoadOptions object
		LoadOptions loadOptions4 = new LoadOptions(FileFormatType.CSV);

		// Creating an Workbook object with CSV file path and the loadOptions
		// object
		new Workbook(dataDir + "Book_CSV.csv", loadOptions4);

		// Print message
		System.out.println("CSV format workbook has been opened successfully.");


	}
}
