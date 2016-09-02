package com.aspose.cells.examples.articles;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class RenderCustomDateFormat {
	public static void main(String[] args) throws Exception {
		// ExStart:RenderCustomDateFormat
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(RenderCustomDateFormat.class);

		Workbook workbook = new Workbook(dataDir + "DateFormat.xlsx");
		workbook.save(dataDir + "out.pdf");
		// ExEnd:RenderCustomDateFormat
	}
}
