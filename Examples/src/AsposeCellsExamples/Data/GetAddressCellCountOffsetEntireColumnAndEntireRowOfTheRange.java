package AsposeCellsExamples.Data;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class GetAddressCellCountOffsetEntireColumnAndEntireRowOfTheRange { 
	
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());

		// Create empty workbook.
		Workbook wb = new Workbook();

		// Access first worksheet.
		Worksheet ws = wb.getWorksheets().get(0);

		// Create range A1:B3.
		System.out.println("Creating Range A1:B3\n");
		Range rng = ws.getCells().createRange("A1:B3");

		// Print range address and cell count.
		System.out.println("Range Address: " + rng.getAddress());
		System.out.println("Cell Count: " + rng.getCellCount());

		// Formatting console output.
		System.out.println("----------------------");
		System.out.println("");

		// Create range A1.
		System.out.println("Creating Range A1\n");
		rng = ws.getCells().createRange("A1");

		// Print range offset, entire column and entire row.
		System.out.println("Offset: " + rng.getOffset(2, 2).getAddress());
		System.out.println("Entire Column: " + rng.getEntireColumn().getAddress());
		System.out.println("Entire Row: " + rng.getEntireRow().getAddress());

		// Formatting console output.
		System.out.println("----------------------");
		System.out.println("");
		 
		// Print the message
		System.out.println("GetAddressCellCountOffsetEntireColumnAndEntireRowOfTheRange executed successfully.");
	}
}
