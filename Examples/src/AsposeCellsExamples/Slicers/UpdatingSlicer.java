package AsposeCellsExamples.Slicers;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class UpdatingSlicer { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		// Load sample Excel file containing slicer.
		Workbook wb = new Workbook(srcDir + "sampleUpdatingSlicer.xlsx");

		// Access first worksheet.
		Worksheet ws = wb.getWorksheets().get(0);

		// Access the first slicer inside the slicer collection.
		Slicer slicer = ws.getSlicers().get(0);

		// Access the slicer items.
		SlicerCacheItemCollection scItems = slicer.getSlicerCache().getSlicerCacheItems();

		// Unselect 2nd and 3rd slicer items.
		scItems.get(1).setSelected(false);
		scItems.get(2).setSelected(false);

		// Refresh the slicer.
		slicer.refresh();

		// Save the workbook in output XLSX format.
		wb.save(outDir + "outputUpdatingSlicer.xlsx", SaveFormat.XLSX);
		
		// Print the message
		System.out.println("UpdatingSlicer executed successfully.");
	}
}
