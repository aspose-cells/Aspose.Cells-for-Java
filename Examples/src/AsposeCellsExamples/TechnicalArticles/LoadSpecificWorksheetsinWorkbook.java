package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.LoadFormat;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class LoadSpecificWorksheetsinWorkbook {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(LoadSpecificWorksheetsinWorkbook.class) + "TechnicalArticles/";

		//Define a new Workbook
		Workbook workbook;

		/// Load the workbook with the specified worksheet only.
		LoadOptions loadOptions = new LoadOptions(LoadFormat.XLSX);
		loadOptions.setLoadFilter(new CustomLoad());

		// Creat the workbook.
		workbook = new Workbook(dataDir+ "TestData.xlsx", loadOptions);

		// Perform your desired task.

		// Save the workbook.
		workbook.save(dataDir+ "outputFile.out.xlsx");

	}
}
