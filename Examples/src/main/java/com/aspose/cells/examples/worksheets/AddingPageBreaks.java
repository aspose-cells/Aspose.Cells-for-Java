package com.aspose.cells.examples.worksheets;

import com.aspose.cells.HorizontalPageBreakCollection;
import com.aspose.cells.VerticalPageBreakCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class AddingPageBreaks {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(AddingPageBreaks.class) + "worksheets/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Add a page break at cell Y30
		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet worksheet = worksheets.get(0);
		HorizontalPageBreakCollection hPageBreaks = worksheet.getHorizontalPageBreaks();
		hPageBreaks.add("Y30");
		VerticalPageBreakCollection vPageBreaks = worksheet.getVerticalPageBreaks();
		vPageBreaks.add("Y30");
	}
}
