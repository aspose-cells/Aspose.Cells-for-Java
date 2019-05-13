package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.LoadDataFilterOptions;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class LoadSourceExcelFile {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(LoadSourceExcelFile.class) + "TechnicalArticles/";
		// Specify the load options and filter the data to not load charts
		LoadOptions options = new LoadOptions();
		options.getLoadFilter().setLoadDataFilterOptions(LoadDataFilterOptions.ALL & ~LoadDataFilterOptions.CHART);
//		options.setLoadDataFilterOptions(LoadDataFilterOptions.ALL & ~LoadDataFilterOptions.CHART);

		// Load the workbook with specified load options
		Workbook workbook = new Workbook(dataDir + "LoadSourceExcelFile.xlsx", options);

		// Save the workbook in output format
		workbook.save(dataDir + "LoadSourceExcelFile_out.pdf", SaveFormat.PDF);
        // ExEnd:1

		System.out.println("LoadSourceExcelFile executed successfully.");
	}

}