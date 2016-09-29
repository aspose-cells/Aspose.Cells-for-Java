package com.aspose.cells.examples.articles;

import com.aspose.cells.CellArea;
import com.aspose.cells.Color;
import com.aspose.cells.FormatCondition;
import com.aspose.cells.FormatConditionCollection;
import com.aspose.cells.FormatConditionType;
import com.aspose.cells.OperatorType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ConditionalFormattingOnCellValue {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ConditionalFormattingOnCellValue.class) + "articles/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();
		Worksheet sheet = workbook.getWorksheets().get(0);
		// Adds an empty conditional formatting
		int index = sheet.getConditionalFormattings().add();
		FormatConditionCollection fcs = sheet.getConditionalFormattings().get(index);

		// Sets the conditional format range.
		CellArea ca = new CellArea();
		ca.StartRow = 0;
		ca.EndRow = 0;
		ca.StartColumn = 0;
		ca.EndColumn = 0;
		fcs.addArea(ca);
		// Sets condition formulas.
		int conditionIndex = fcs.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "50", "100");
		FormatCondition fc = fcs.get(conditionIndex);
		fc.getStyle().setBackgroundColor(Color.getRed());
		workbook.save(dataDir + "CFOnCellValue_out.xls");

	}
}
