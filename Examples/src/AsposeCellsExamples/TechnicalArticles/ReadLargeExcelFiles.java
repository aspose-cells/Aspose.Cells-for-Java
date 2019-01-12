package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.LoadOptions;
import com.aspose.cells.MemorySetting;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class ReadLargeExcelFiles {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ReadLargeExcelFiles.class) + "TechnicalArticles/";
		// Specify the LoadOptions
		LoadOptions opt = new LoadOptions();
		// Set the memory preferences
		opt.setMemorySetting(MemorySetting.MEMORY_PREFERENCE);
		// Instantiate the Workbook
		// Load the Big Excel file having large Data set in it
		new Workbook(dataDir + "RLExcelFiles_out.xlsx", opt);
	}
}
