package com.aspose.cells.examples.worksheets.Value;

import com.aspose.cells.HorizontalPageBreakCollection;
import com.aspose.cells.VerticalPageBreakCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class RemoveSpecificPageBreak {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getDataDir(RemoveSpecificPageBreak.class);
		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Removing a specific page break
		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet worksheet = worksheets.get(0);
		HorizontalPageBreakCollection hPageBreaks = worksheet.getHorizontalPageBreaks();
		hPageBreaks.removeAt(0);
		VerticalPageBreakCollection vPageBreaks = worksheet.getVerticalPageBreaks();
		vPageBreaks.removeAt(0);
	}
}
