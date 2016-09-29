package com.aspose.cells.examples.articles;

import com.aspose.cells.CellArea;
import com.aspose.cells.Color;
import com.aspose.cells.ConditionalFormattingCollection;
import com.aspose.cells.FormatCondition;
import com.aspose.cells.FormatConditionCollection;
import com.aspose.cells.FormatConditionType;
import com.aspose.cells.OperatorType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ConditionalFormattingBasedOnFormula {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ConditionalFormattingBasedOnFormula.class) + "articles/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();
		Worksheet sheet = workbook.getWorksheets().get(0);
		ConditionalFormattingCollection cfs = sheet.getConditionalFormattings();
		int index = cfs.add();
		FormatConditionCollection fcs = cfs.get(index);
		// Sets the conditional format range.
		CellArea ca = new CellArea();
		ca = new CellArea();
		ca.StartRow = 2;
		ca.EndRow = 2;
		ca.StartColumn = 1;
		ca.EndColumn = 1;
		fcs.addArea(ca);
		// Sets condition formulas.
		int conditionIndex = fcs.addCondition(FormatConditionType.EXPRESSION, OperatorType.NONE, "", "");
		FormatCondition fc = fcs.get(conditionIndex);
		fc.setFormula1("=IF(SUM(B1:B2)>100,TRUE,FALSE)");
		fc.getStyle().setBackgroundColor(Color.getRed());
		sheet.getCells().get("B3").setFormula("=SUM(B1:B2)");
		sheet.getCells().get("C4").setValue("If Sum of B1:B2 is greater than 100, B3 will have RED background");
		workbook.save(dataDir + "CFBasedOnFormula_out.xls");

	}
}
