package AsposeCellsExamples.Slicers;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class FormattingSlicer {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());

		// Load sample Excel file containing slicer.
		Workbook wb = new Workbook(srcDir + "sampleFormattingSlicer.xlsx");

		// Access first worksheet.
		Worksheet ws = wb.getWorksheets().get(0);

		// Access the first slicer inside the slicer collection.
		Slicer slicer = ws.getSlicers().get(0);

		// Set the number of columns of the slicer.
		slicer.setNumberOfColumns(2);

		// Set the type of slicer style.
		slicer.setStyleType(SlicerStyleType.SLICER_STYLE_LIGHT_6);

		// Save the workbook in output XLSX format.
		wb.save(outDir + "outputFormattingSlicer.xlsx", SaveFormat.XLSX);
		 
		// Print the message
		System.out.println("FormattingSlicer executed successfully.");
	}
}
