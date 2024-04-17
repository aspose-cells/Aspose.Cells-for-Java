package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.SparklineGroup;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class CopySparkline {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CopySparkline.class) + "TechnicalArticles/";
		// Create workbook from source Excel file
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access the first sparkline group
		SparklineGroup group = worksheet.getSparklineGroups().get(0);

		// Add Data Ranges and Locations inside this sparkline group
		group.getSparklines().add("D5:O5", 4, 15);
		group.getSparklines().add("D6:O6", 5, 15);
		group.getSparklines().add("D7:O7", 6, 15);
		group.getSparklines().add("D8:O8", 7, 15);

		// Save the workbook
		workbook.save(dataDir + "CopySparkline_out.xlsx");

	}
}
