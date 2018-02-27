package AsposeCellsExamples.Formulas;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class SpecifyMaximumRowsOfSharedFormula { 
	
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
				
		//Create workbook
		Workbook wb = new Workbook();

		//Set the max rows of shared formula to 5
		wb.getSettings().setMaxRowsOfSharedFormula(5);

		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		//Access cell D1
		Cell cell = ws.getCells().get("D1");

		//Set the shared formula in 100 rows
		cell.setSharedFormula("=Sum(A1:A2)", 100, 1);

		//Save the output Excel file
		wb.save(outDir + "outputSpecifyMaximumRowsOfSharedFormula.xlsx");

		// Print the message
		System.out.println("SpecifyMaximumRowsOfSharedFormula executed successfully.");
	}
}
