package com.aspose.cells.examples.worksheets.Value;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ClearAllPageBreaks {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getDataDir(ClearAllPageBreaks.class);
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();
		workbook.getWorksheets().get(0).getHorizontalPageBreaks().clear();
		workbook.getWorksheets().get(0).getVerticalPageBreaks().clear();
	}
}
