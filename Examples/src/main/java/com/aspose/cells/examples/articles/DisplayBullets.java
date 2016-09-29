package com.aspose.cells.examples.articles;

import com.aspose.cells.Cell;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class DisplayBullets {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(DisplayBullets.class) + "articles/";
		//Create workbook object
		Workbook workbook = new Workbook();

		//Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		//Access cell A1
		Cell cell = worksheet.getCells().get("A1");

		//Set the HTML string
		cell.setHtmlString("<font style='font-family:Arial;font-size:10pt;color:#666666;vertical-align:top;text-align:left;'>Text 1 </font><font style='font-family:Wingdings;font-size:8.0pt;color:#009DD9;mso-font-charset:2;'>l</font><font style='font-family:Arial;font-size:10pt;color:#666666;vertical-align:top;text-align:left;'> Text 2 </font><font style='font-family:Wingdings;font-size:8.0pt;color:#009DD9;mso-font-charset:2;'>l</font><font style='font-family:Arial;font-size:10pt;color:#666666;vertical-align:top;text-align:left;'> Text 3 </font><font style='font-family:Wingdings;font-size:8.0pt;color:#009DD9;mso-font-charset:2;'>l</font><font style='font-family:Arial;font-size:10pt;color:#666666;vertical-align:top;text-align:left;'> Text 4 </font>");

		//Save the workbook
		workbook.save(dataDir + "DisplayBullets_out.xlsx");

	}
}
