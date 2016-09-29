package com.aspose.cells.examples.articles;

import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ExportExceltoHTML {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ExportExceltoHTML.class) + "articles/";
		// Create your workbook
		Workbook wb = new Workbook();

		// Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		// Fill worksheet with some integer values
		for (int r = 0; r < 10; r++) {
			for (int c = 0; c < 10; c++) {
				ws.getCells().get(r, c).putValue(r * 1);
			}
		}

		// Save your workbook in HTML format and export gridlines
		HtmlSaveOptions opts = new HtmlSaveOptions();
		opts.setExportGridLines(true);
		wb.save(dataDir + "EExceltoHTML_out.html", opts);

	}

}
