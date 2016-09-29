package com.aspose.cells.examples.articles;

import java.awt.image.BufferedImage;
import java.io.File;

import javax.imageio.ImageIO;

import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;
import com.sun.prism.Image;

public class GenerateThumbnailofWorksheet {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(GenerateThumbnailofWorksheet.class) + "articles/";
		// Instantiate and open an Excel file
		Workbook book = new Workbook(dataDir + "book1.xls");

		// Define ImageOrPrintOptions
		ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
		// Set the vertical and horizontal resolution
		imgOptions.setVerticalResolution(200);
		imgOptions.setHorizontalResolution(200);
		// Set the image's format
		imgOptions.setImageFormat(ImageFormat.getJpeg());
		// One page per sheet is enabled
		imgOptions.setOnePagePerSheet(true);

		// Get the first worksheet
		Worksheet sheet = book.getWorksheets().get(0);
		// Render the sheet with respect to specified image/print options
		SheetRender sr = new SheetRender(sheet, imgOptions);
		// Render the image for the sheet
		sr.toImage(0, "mythumb.jpg");

		// Creating Thumbnail
		java.awt.Image img = ImageIO.read(new File(dataDir + "school.jpg")).getScaledInstance(100, 100, BufferedImage.SCALE_SMOOTH);
		BufferedImage img1 = new BufferedImage(100, 100, BufferedImage.TYPE_INT_RGB);
		img1.createGraphics().drawImage(
				ImageIO.read(new File(dataDir + "school.jpg")).getScaledInstance(100, 100, img.SCALE_SMOOTH), 0, 0, null);
		ImageIO.write(img1, "jpg", new File(dataDir + "GTOfWorksheet_out.jpg"));

	}
}
