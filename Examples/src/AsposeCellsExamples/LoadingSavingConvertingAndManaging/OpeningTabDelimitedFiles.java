package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class OpeningTabDelimitedFiles {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(OpeningTabDelimitedFiles.class) + "LoadingSavingConvertingAndManaging/";

		// Creating and TAB_DELIMITED LoadOptions object
		LoadOptions loadOptions5 = new LoadOptions(FileFormatType.TAB_DELIMITED);

		// Creating an Workbook object with Tab Delimited text file path and the
		// loadOptions object
		Workbook workbook7 = new Workbook(dataDir + "Book1TabDelimited.txt", loadOptions5);

		System.out.println(workbook7.getFileName());
		// Print message
		System.out.println("Tab Delimited workbook has been opened successfully.");


	}
}
