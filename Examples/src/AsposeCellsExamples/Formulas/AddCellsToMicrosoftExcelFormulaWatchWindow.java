package AsposeCellsExamples.Formulas;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class AddCellsToMicrosoftExcelFormulaWatchWindow {
	
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		// Create empty workbook.
		Workbook wb = new Workbook();

		// Access first worksheet.
		Worksheet ws = wb.getWorksheets().get(0);

		// Put some integer values in cell A1 and A2.
		ws.getCells().get("A1").putValue(10);
		ws.getCells().get("A2").putValue(30);

		// Access cell C1 and set its formula.
		Cell c1 = ws.getCells().get("C1");
		c1.setFormula("=Sum(A1,A2)");

		// Add cell C1 into cell watches by name.
		ws.getCellWatches().add(c1.getName());

		// Access cell E1 and set its formula.
		Cell e1 = ws.getCells().get("E1");
		e1.setFormula("=A2*A1");

		// Add cell E1 into cell watches by its row and column indices.
		ws.getCellWatches().add(e1.getRow(), e1.getColumn());

		// Save workbook to output XLSX format.
		wb.save(outDir + "outputAddCellsToMicrosoftExcelFormulaWatchWindow.xlsx", SaveFormat.XLSX);
		
		// Print the message
		System.out.println("AddCellsToMicrosoftExcelFormulaWatchWindow executed successfully.");
	}
}

