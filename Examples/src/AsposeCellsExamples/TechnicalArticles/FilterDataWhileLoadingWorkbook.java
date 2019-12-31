package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.LoadDataFilterOptions;
import com.aspose.cells.LoadFormat;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class FilterDataWhileLoadingWorkbook {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(FilterDataWhileLoadingWorkbook.class) + "TechnicalArticles/";
		// Set the load options, we only want to load shapes and do not want to load data
		LoadOptions opts = new LoadOptions(LoadFormat.XLSX);
		opts.getLoadFilter().setLoadDataFilterOptions(LoadDataFilterOptions.DRAWING);

		// Create workbook object from sample excel file using load options
		Workbook wb = new Workbook(dataDir + "sampleFilterDataWhileLoadingWorkbook.xlsx", opts);

		// Save the output in PDF format
		wb.save(dataDir + "sampleFilterDataWhileLoadingWorkbook_out.pdf", SaveFormat.PDF);
        // ExEnd:1

		System.out.println("FilterDataWhileLoadingWorkbook executed successfully.");

	}

}
