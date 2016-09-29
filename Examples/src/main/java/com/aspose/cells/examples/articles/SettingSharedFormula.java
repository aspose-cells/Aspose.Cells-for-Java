package com.aspose.cells.examples.articles;

import com.aspose.cells.Cells;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class SettingSharedFormula {
	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(SettingSharedFormula.class) + "articles/";

		String filePath = dataDir + "input.xlsx";

		// Instantiate a Workbook from existing file
		Workbook workbook = new Workbook(filePath);

		// Get the cells collection in the first worksheet
		Cells cells = workbook.getWorksheets().get(0).getCells();

		// Apply the shared formula in the range i.e.., B2:B14
		cells.get("B2").setSharedFormula("=A2*0.09", 13, 1);

		// Save the excel file
		workbook.save(dataDir + "SSharedFormula_out.xlsx", SaveFormat.XLSX);

	}
}
