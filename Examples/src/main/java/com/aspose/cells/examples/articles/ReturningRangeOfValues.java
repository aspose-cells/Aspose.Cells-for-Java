package com.aspose.cells.examples.articles;

import com.aspose.cells.CalcModeType;
import com.aspose.cells.CalculationOptions;
import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ReturningRangeOfValues {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ReturningRangeOfValues.class) + "articles/";
		Workbook wb = new Workbook();
		Cells cells = wb.getWorksheets().get(0).getCells();

		Cell cell = cells.get(0, 0);
		cell.setArrayFormula("=MYFUNC()", 2, 2);

		Style style = cell.getStyle();
		style.setNumber(14);
		cell.setStyle(style);

		CalculationOptions copt = new CalculationOptions();
		copt.setCustomFunction(new CustomFunctionStaticValue());
		wb.calculateFormula(copt);

		// Save to xlsx by setting the calc mode to manual
		wb.getSettings().setCalcMode(CalcModeType.MANUAL);
		wb.save(dataDir + "output.xlsx");

		// Save to pdf
		wb.save(dataDir + "output.pdf");


	}
}
