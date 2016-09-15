package com.aspose.cells.examples.articles;

import com.aspose.cells.AbstractCalculationEngine;
import com.aspose.cells.CalculationData;
import com.aspose.cells.CalculationOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class CalculateCustomFunctionWithoutWritingToWorksheet {

	public class CustomEngine extends AbstractCalculationEngine {
		public void calculate(CalculationData data) {

			// Check the forumla name and calculate it yourself
			if (data.getFunctionName().equals("MyCompany.CustomFunction")) {
				// This is our calculated value
				data.setCalculatedValue("Aspose.Cells.");
			}

		}
	}

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