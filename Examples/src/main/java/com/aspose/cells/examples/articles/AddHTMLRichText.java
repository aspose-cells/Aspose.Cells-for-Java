package com.aspose.cells.examples.articles;

import com.aspose.cells.Cell;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AddHTMLRichText {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddHTMLRichText.class) + "articles/";
		Workbook workbook = new Workbook();

		Worksheet worksheet = workbook.getWorksheets().get(0);

		Cell cell = worksheet.getCells().get("A1");
		cell.setHtmlString(
				"<Font Style=\"FONT-WEIGHT: bold;FONT-STYLE: italic;TEXT-DECORATION: underline;FONT-FAMILY: Arial;FONT-SIZE: 11pt;COLOR: #ff0000;\">This is simple HTML formatted text.</Font>");
		workbook.save(dataDir + "AHTMLRText_out.xlsx");

	}
}
