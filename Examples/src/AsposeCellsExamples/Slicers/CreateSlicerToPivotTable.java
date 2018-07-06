package AsposeCellsExamples.Slicers;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class CreateSlicerToPivotTable {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		// Load sample Excel file containing pivot table.
		Workbook wb = new Workbook(srcDir + "sampleCreateSlicerToPivotTable.xlsx");
		 
		// Access first worksheet.
		Worksheet ws = wb.getWorksheets().get(0);
		 
		// Access first pivot table inside the worksheet.
		PivotTable pt = ws.getPivotTables().get(0);
		 
		// Add slicer relating to pivot table with first base field at cell B22.
		int idx = ws.getSlicers().add(pt, "B22", pt.getBaseFields().get(0));
		 
		// Access the newly added slicer from slicer collection.
		Slicer slicer = ws.getSlicers().get(idx);
		 
		// Save the workbook in output XLSX format.
		wb.save(outDir + "outputCreateSlicerToPivotTable.xlsx", SaveFormat.XLSX);
		 
		// Save the workbook in output XLSB format.
		wb.save(outDir + "outputCreateSlicerToPivotTable.xlsb", SaveFormat.XLSB);

		// Print the message
		System.out.println("CreateSlicerToPivotTable executed successfully.");
	}
}
