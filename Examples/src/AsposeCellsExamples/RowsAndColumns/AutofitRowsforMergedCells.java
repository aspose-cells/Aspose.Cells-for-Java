package AsposeCellsExamples.RowsAndColumns;

import com.aspose.cells.AutoFitMergedCellsType;
import com.aspose.cells.AutoFitterOptions;
import com.aspose.cells.Range;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class AutofitRowsforMergedCells {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AutofitRowsforMergedCells.class) + "RowsAndColumns/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);
		
		// Create a range A1:B1
		Range range = worksheet.getCells().createRange(0, 0, 1, 2);
		
		// Merge the cells
		range.merge();
		
		// Insert value to the merged cell A1
		worksheet.getCells().get(0, 0).setValue("A quick brown fox jumps over the lazy dog. A quick brown fox jumps over the lazy dog....end");
		
		// Create a style object
		Style style = worksheet.getCells().get(0, 0).getStyle();

		// Set wrapping text on
		style.setTextWrapped(true);

		// Apply the style to the cell
		worksheet.getCells().get(0, 0).setStyle(style);

		// Create an object for AutoFitterOptions
		AutoFitterOptions options = new AutoFitterOptions();

		// Set auto-fit for merged cells
		options.setAutoFitMergedCellsType(AutoFitMergedCellsType.EACH_LINE);

		// Autofit rows in the sheet(including the merged cells)
		worksheet.autoFitRows(options);

		// Save the Excel file
		workbook.save(dataDir + "AutofitRowsforMergedCells_out.xlsx");
		// ExEnd:1
		        
		System.out.println("AutofitRowsforMergedCells executed successfully.");
	}
}
