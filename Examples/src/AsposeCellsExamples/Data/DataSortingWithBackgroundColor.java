package AsposeCellsExamples.Data;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

public class DataSortingWithBackgroundColor {

	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// Load the source Excel file
		Workbook workbook = new Workbook(srcDir + "sampleBackgroundFile.xlsx");

		// Instantiate data sorter object
		DataSorter sorter = workbook.getDataSorter();

		// Add key for Column B, Sort it in descending order with background color red
		sorter.addKey(1, SortOnType.CELL_COLOR, SortOrder.DESCENDING, Color.getRed());

		// Sort the data based on the key
		sorter.sort(workbook.getWorksheets().get(0).getCells(), CellArea.createCellArea("A2", "C6"));

		// Save the output file
		workbook.save(outDir + "outputSampleBackgroundFile.xlsx");
		// ExEnd:1

		// Print message
		System.out.println("DataSortingWithBackgroundColor Executed Successfully");
	}
}
