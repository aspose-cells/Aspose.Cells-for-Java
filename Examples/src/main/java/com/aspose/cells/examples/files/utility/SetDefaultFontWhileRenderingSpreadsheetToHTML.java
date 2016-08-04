
package com.aspose.cells.examples.files.utility;

import com.aspose.cells.Cell;
import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SetDefaultFontWhileRenderingSpreadsheetToHTML {

	public static void main(String[] args) throws Exception {
		
		// Directory path where output HTML files are to be saved
		String dataDir = Utils.getSharedDataDir(SetDefaultFontWhileRenderingSpreadsheetToHTML.class) + "Conversion/";
		
		//Create workbook object.
		Workbook wb = new Workbook();

		//Access first WorkSheet.
		Worksheet ws = wb.getWorksheets().get(0);
		
		//Access cell B4 and add some text inside it.
		Cell cell = ws.getCells().get("B4");
		cell.putValue("This text has some unknown or invalid font which does not exist.");

		//Set the font of cell B4 which is unknown.
		Style st = cell.getStyle();
		st.getFont().setName("UnknownNotExist");
		st.getFont().setSize(20);
		cell.setStyle(st);

		//Now save the workbook in html format and set the default font to Courier New.
		HtmlSaveOptions opts = new HtmlSaveOptions();
		opts.setDefaultFontName("Courier New");
		wb.save(dataDir + "out_courier_new.htm", opts);

		//Now save the workbook in html format once again but set the default font to Arial.
		opts.setDefaultFontName("Arial");
		wb.save(dataDir + "out_arial.htm", opts);

		//Now save the workbook in html format once again but set the default font to Times New Roman.
		opts.setDefaultFontName("Times New Roman");
		wb.save(dataDir + "out_times_new_roman.htm", opts);

	}

}
