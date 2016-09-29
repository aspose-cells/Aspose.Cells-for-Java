package com.aspose.cells.examples.articles;

import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class CreateTransparentImage {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CreateTransparentImage.class) + "articles/";
		// Create workbook object from source file
		Workbook wb = new Workbook(dataDir + "aspose-sample.xlsx");

		// Apply different image or print options
		ImageOrPrintOptions imgOption = new ImageOrPrintOptions();
		imgOption.setImageFormat(ImageFormat.getPng());
		imgOption.setHorizontalResolution(200);
		imgOption.setVerticalResolution(200);
		imgOption.setOnePagePerSheet(true);

		// Apply transparency to the output image
		imgOption.setTransparent(true);

		// Create image after apply image or print options
		SheetRender sr = new SheetRender(wb.getWorksheets().get(0), imgOption);
		sr.toImage(0, dataDir + "CTransparentImage_out.png");

	}
}
