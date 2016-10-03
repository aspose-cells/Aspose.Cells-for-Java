package com.aspose.cells.examples.data;

import com.aspose.cells.CellArea;
import com.aspose.cells.ConditionalFormattingCollection;
import com.aspose.cells.FormatCondition;
import com.aspose.cells.FormatConditionCollection;
import com.aspose.cells.FormatConditionType;
import com.aspose.cells.OperatorType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ConditionalFormattingatRuntime {
	public static void main(String[] args) throws Exception {
		// Path to source file
		String dataDir = Utils.getSharedDataDir(ConditionalFormattingatRuntime.class) + "data/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();
		Worksheet sheet = workbook.getWorksheets().get(0);
		ConditionalFormattingCollection cfs = sheet.getConditionalFormattings();

		// The first method:adds an empty conditional formatting
		int index = cfs.add();
		FormatConditionCollection fcs = cfs.get(index);

		// Sets the conditional format range.
		CellArea ca1 = new CellArea();
		ca1.StartRow = 0;
		ca1.StartColumn = 0;
		ca1.EndRow = 0;
		ca1.EndColumn = 0;

		CellArea ca2 = new CellArea();
		ca2.StartRow = 0;
		ca2.StartColumn = 0;
		ca2.EndRow = 0;
		ca2.EndColumn = 0;

		CellArea ca3 = new CellArea();
		ca3.StartRow = 0;
		ca3.StartColumn = 0;
		ca3.EndRow = 0;
		ca3.EndColumn = 0;

		fcs.addArea(ca1);
		fcs.addArea(ca2);
		fcs.addArea(ca3);

		// Sets condition formulas.
		int conditionIndex = fcs.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "=A2", "100");

		FormatCondition fc = fcs.get(conditionIndex);

		int conditionIndex2 = fcs.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "50", "100");

		// Saving the Excel file
		workbook.save(dataDir + "CFAtRuntime_out.xls");
	}
}
