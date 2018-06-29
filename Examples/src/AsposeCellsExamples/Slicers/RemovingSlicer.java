package AsposeCellsExamples.Slicers;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class RemovingSlicer {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		// Load sample Excel file containing slicer.
		Workbook wb = new Workbook(srcDir + "sampleRemovingSlicer.xlsx");

		// Access first worksheet.
		Worksheet ws = wb.getWorksheets().get(0);

		// Access the first slicer inside the slicer collection.
		Slicer slicer = ws.getSlicers().get(0);

		// Remove slicer.
		ws.getSlicers().remove(slicer);

		// Save the workbook in output XLSX format.
		wb.save(outDir + "outputRemovingSlicer.xlsx", SaveFormat.XLSX);
		
		// Print the message
		System.out.println("RemovingSlicer executed successfully.");
	}
}
