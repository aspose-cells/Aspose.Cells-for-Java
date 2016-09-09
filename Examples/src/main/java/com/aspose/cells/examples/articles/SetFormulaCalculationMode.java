package com.aspose.cells.examples.articles;

import com.aspose.cells.CalcModeType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class SetFormulaCalculationMode {
	public static void main(String[] args) throws Exception {
		// ExStart:SetFormulaCalculationMode
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SetFormulaCalculationMode.class) + "articles/";
		// Create a workbook
		Workbook workbook = new Workbook();

		// Set the Formula Calculation Mode to Manual
		workbook.getSettings().setCalcMode(CalcModeType.MANUAL);

		// Save the workbook
		workbook.save("SFCalculationMode-out.xlsx", SaveFormat.XLSX);
		// ExEnd:SetFormulaCalculationMode
	}
}
