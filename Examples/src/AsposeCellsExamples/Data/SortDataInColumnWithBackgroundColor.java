package AsposeCellsExamples.Data;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class SortDataInColumnWithBackgroundColor {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

        // ExStart:1
        // Load the source Excel file
        Workbook workbook = new Workbook(srcDir + "CellsNet46500.xlsx");

        // Instantiate data sorter object
        DataSorter sorter = workbook.getDataSorter();

        // Add key for Column B, Sort it in descending order with background color red
        sorter.addKey(1, SortOnType.CELL_COLOR, SortOrder.DESCENDING, Color.getRed());

        // Sort the data based on the key
        sorter.sort(workbook.getWorksheets().get(0).getCells(), CellArea.createCellArea("A2", "C6"));

        // Save the output file
        workbook.save(outDir + "outputSortData_CustomSortList.xlsx");
        // ExEnd:1

        // Print the message
		System.out.println("SortDataInColumnWithCustomSortList executed successfully.");
	}
}
