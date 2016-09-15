package com.aspose.cells.examples.articles;

import com.aspose.cells.AbstractCalculationEngine;
import com.aspose.cells.CalculationData;
import com.aspose.cells.CalculationOptions;
import com.aspose.cells.Cell;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ImplementCustomCalculationEngine {


	public class CustomEngine extends AbstractCalculationEngine {
		public void calculate(CalculationData data) {
			if (data.getFunctionName().toUpperCase().equals("SUM") == true) {
				double val = (double) data.getCalculatedValue();
				val = val + 30;

				data.setCalculatedValue(val);
			}
		}
	}

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getDataDir(ImplementCustomCalculationEngine.class);

		Workbook workbook = new Workbook();

		Worksheet sheet = workbook.getWorksheets().get(0);

		Cell a1 = sheet.getCells().get("A1");
		a1.setFormula("=Sum(B1:B2)");

		sheet.getCells().get("B1").putValue(10);
		sheet.getCells().get("B2").putValue(10);

		// Without Custom Engine, the value of cell A1 will be 20
		workbook.calculateFormula();

		System.out.println("Without Custom Engine Value of A1: " + a1.getStringValue());

		// With Custom Engine, the value of cell A1 will be 50

		CalculationOptions opts = new CalculationOptions();
		opts.setCustomEngine(new CustomEngine());

		workbook.calculateFormula(opts);

		System.out.println("With Custom Engine Value of A1: " + a1.getStringValue());

	}


}
