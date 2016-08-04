package com.aspose.cells.examples.files.utility;

import com.aspose.cells.Cell;
import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SetDefaultFontWhileRenderingSpreadsheetToImages {

	public static void main(String[] args) throws Exception {
		
		// Directory path where output HTML files are to be saved
		String dataDir = Utils.getSharedDataDir(SetDefaultFontWhileRenderingSpreadsheetToHTML.class) + "Conversion/";
		
		//Create workbook object.
		Workbook wb = new Workbook();

		//Set default font of the workbook to none
		Style s = wb.getDefaultStyle();
		s.getFont().setName("");
		wb.setDefaultStyle(s);

		//Access first worksheet.
		Worksheet ws = wb.getWorksheets().get(0);

		//Access cell A4 and add some text inside it.
		Cell cell = ws.getCells().get("A4");
		cell.putValue("This text has some unknown or invalid font which does not exist.");

		//Set the font of cell A4 which is unknown.
		Style st = cell.getStyle();
		st.getFont().setName("UnknownNotExist");
		st.getFont().setSize(20);
		st.setTextWrapped(true);
		cell.setStyle(st);

		//Set first column width and fourth column height
		ws.getCells().setColumnWidth(0, 80);
		ws.getCells().setRowHeight(3, 60);

		//Create image or print options.
		ImageOrPrintOptions opts = new ImageOrPrintOptions();
		opts.setOnePagePerSheet(true);
		opts.setImageFormat(ImageFormat.getPng());

		//Render worksheet image with Courier New as default font.
		opts.setDefaultFont("Courier New");
		SheetRender sr = new SheetRender(ws, opts);
		sr.toImage(0, dataDir + "out_courier_new.png");

		//Render worksheet image again with Times New Roman as default font.
		opts.setDefaultFont("Times New Roman");
		sr = new SheetRender(ws, opts);
		sr.toImage(0, dataDir + "out_times_new_roman.png");

	}

}
