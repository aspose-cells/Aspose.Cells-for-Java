package com.aspose.cells.examples.articles;

import com.aspose.cells.CalculationOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class DecreaseCalculationTime {

	public static void main(String[] args) throws Exception {

		// Test calculation time after setting recursive true
		TestCalcTimeRecursive(true);

		// Test calculation time after setting recursive false
		TestCalcTimeRecursive(false);
	}

	// --------------------------------------------------

	static void TestCalcTimeRecursive(boolean rec) throws Exception {

		String dataDir = Utils.getSharedDataDir(DecreaseCalculationTime.class) + "articles/";

		// Load your sample workbook
		Workbook wb = new Workbook(dataDir + "sample.xlsx");

		// Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		// Set the calculation option, set recursive true or false as per parameter
		CalculationOptions opts = new CalculationOptions();
		opts.setRecursive(rec);

		// Start calculating time in nanoseconds
		long startTime = System.nanoTime();

		// Calculate cell A1 one million times
		for (int i = 0; i < 1000000; i++) {
			ws.getCells().get("A1").calculate(opts);
		}

		// Calculate elapsed time in seconds
		long second = 1000000000;
		long estimatedTime = System.nanoTime() - startTime;
		estimatedTime = estimatedTime / second;

		// Print the elapsed time in seconds
		System.out.println("Recursive " + rec + ": " + estimatedTime + " seconds");
	}

}
