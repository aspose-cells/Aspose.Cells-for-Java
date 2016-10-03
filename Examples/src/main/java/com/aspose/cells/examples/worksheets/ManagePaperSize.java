package com.aspose.cells.examples.worksheets;

import com.aspose.cells.PageSetup;
import com.aspose.cells.PaperSizeType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class ManagePaperSize {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(ManagePaperSize.class) + "worksheets/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Accessing the first worksheet in the Excel file
		WorksheetCollection worksheets = workbook.getWorksheets();
		int sheetIndex = worksheets.add();
		Worksheet sheet = worksheets.get(sheetIndex);

		// Setting the paper size to A4
		PageSetup pageSetup = sheet.getPageSetup();
		pageSetup.setPaperSize(PaperSizeType.PAPER_A_4);

		workbook.save(dataDir + "ManagePaperSize_out.xls");
	}
}
