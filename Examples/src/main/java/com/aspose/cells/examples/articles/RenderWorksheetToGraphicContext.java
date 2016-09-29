package com.aspose.cells.examples.articles;

import java.awt.Graphics2D;
import java.awt.image.BufferedImage;

import javax.imageio.ImageIO;

import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;
import java.io.*;

public class RenderWorksheetToGraphicContext {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(RenderWorksheetToGraphicContext.class) + "articles/";

		// Create workbook object from source file
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Create empty image and fill it with blue color
		int width = 800;
		int height = 800;
		BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
		Graphics2D g = image.createGraphics();
		g.setColor(java.awt.Color.blue);
		g.fillRect(0, 0, width, height);

		// Set one page per sheet to true in image or print options
		ImageOrPrintOptions opts = new ImageOrPrintOptions();
		opts.setOnePagePerSheet(true);

		// Render worksheet to graphics context
		SheetRender sr = new SheetRender(worksheet, opts);
		sr.toImage(0, g);

		// Save the graphics context image in Png format
		File outputfile = new File(dataDir + "RWToGraphicContext_out.png");
		ImageIO.write(image, "png", outputfile);

	}
}
