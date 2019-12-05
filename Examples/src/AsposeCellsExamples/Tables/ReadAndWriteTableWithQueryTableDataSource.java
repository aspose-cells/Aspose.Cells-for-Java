package AsposeCellsExamples.Tables;

import AsposeCellsExamples.Utils;
import com.aspose.cells.ListObject;
import com.aspose.cells.TableDataSourceType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class ReadAndWriteTableWithQueryTableDataSource {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the output directory.
		String sourceDir = Utils.Get_SourceDirectory();
		String outputDir = Utils.Get_OutputDirectory();

		// Load workbook object
		Workbook workbook = new Workbook(sourceDir + "SampleTableWithQueryTable.xls");

		Worksheet worksheet = workbook.getWorksheets().get(0);

		ListObject table = worksheet.getListObjects().get(0);

		// Check the data source type if it is query table
		if (table.getDataSourceType() == TableDataSourceType.QUERY_TABLE)
		{
			table.setShowTotals(true);
		}

		// Save the file
		workbook.save(outputDir + "SampleTableWithQueryTable_out.xls");
		// ExEnd:1

		System.out.println("ReadAndWriteTableWithQueryTableDataSource executed successfully.");
	}
}
