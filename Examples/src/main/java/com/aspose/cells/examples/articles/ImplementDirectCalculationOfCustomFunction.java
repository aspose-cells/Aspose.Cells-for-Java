package com.aspose.cells.examples.articles;

import com.aspose.cells.AbstractCalculationEngine;
import com.aspose.cells.CalculationOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class ImplementDirectCalculationOfCustomFunction {

	public abstract class CalculateCustomFunctionWithoutWritingToWorksheet extends AbstractCalculationEngine {

		public void Run() {
			// TODO Auto-generated method stub
			// Create a workbook
			Workbook wb = new Workbook();

			// Accesss first worksheet
			Worksheet ws = wb.getWorksheets().get(0);

			// Add some text in cell A1
			ws.getCells().get("A1").putValue("Welcome to ");

			// Create a calculation options with custom engine
			CalculationOptions opts = new CalculationOptions();
			opts.setCustomEngine(new CustomEngine());

			// This line shows how you can call your own custom function without
			// a need to write it in any worksheet cell
			// After the execution of this line, it will return
			// Welcome to Aspose.Cells.
			Object ret = ws.calculateFormula("=A1 & MyCompany.CustomFunction()", opts);

			// Print the calculated value on Console
			System.out.println("Calculated Value: " + ret.toString());
		}

	}

	public static void main(String[] args) throws Exception {
		CalculateCustomFunctionWithoutWritingToWorksheet pg = new CalculateCustomFunctionWithoutWritingToWorksheet();
		pg.Run();
	}

}
