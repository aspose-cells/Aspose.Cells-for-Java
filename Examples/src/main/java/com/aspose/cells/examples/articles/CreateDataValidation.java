package com.aspose.cells.examples.articles;

import com.aspose.cells.OperatorType;
import com.aspose.cells.examples.Utils;
import com.aspose.gridweb.GridCell;
import com.aspose.gridweb.GridValidation;
import com.aspose.gridweb.GridValidationType;
import com.aspose.gridweb.GridWorksheet;

public class CreateDataValidation {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CreateDataValidation.class) + "articles/";

		// Access first worksheet
		GridWorksheet sheet = gridweb.getWorkSheets().get(0);

		// Access cell B3
		GridCell cell = sheet.getCells().get("B3");

		/*
		 * Add validation inside the gridcell Any value which is not between 20
		 * and 40 will cause error in a gridcell
		 */
		GridValidation val = cell.createValidation(GridValidationType.WHOLE_NUMBER, true);
		val.setFormula1("=20");
		val.setFormula2("=40");
		val.setOperator(OperatorType.BETWEEN);
		val.setShowError(true);
		val.setShowInput(true);

	}

}
