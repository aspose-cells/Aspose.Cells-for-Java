package AsposeCellsExamples.TechnicalArticles;

import java.awt.image.BufferedImage;
import java.io.File;

import javax.imageio.ImageIO;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class GenerateThumbnailofWorksheet {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(GenerateThumbnailofWorksheet.class) + "TechnicalArticles/";
		// Instantiate and open an Excel file
		Workbook book = new Workbook(dataDir + "book1.xlsx");

		// Define ImageOrPrintOptions
		ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
		// Set the vertical and horizontal resolution
		imgOptions.setVerticalResolution(200);
		imgOptions.setHorizontalResolution(200);
		// Set the image's format
		imgOptions.setImageType(ImageType.JPEG);
		// One page per sheet is enabled
		imgOptions.setOnePagePerSheet(true);

		// Get the first worksheet
		Worksheet sheet = book.getWorksheets().get(0);
		// Render the sheet with respect to specified image/print options
		SheetRender sr = new SheetRender(sheet, imgOptions);
		// Render the image for the sheet
		sr.toImage(0, dataDir + "mythumb.jpg");

		// Creating Thumbnail
		java.awt.Image img = ImageIO.read(new File(dataDir + "mythumb.jpg")).getScaledInstance(100, 100, BufferedImage.SCALE_SMOOTH);
		BufferedImage img1 = new BufferedImage(100, 100, BufferedImage.TYPE_INT_RGB);
		img1.createGraphics().drawImage(
				ImageIO.read(new File(dataDir + "mythumb.jpg")).getScaledInstance(100, 100, img.SCALE_SMOOTH), 0, 0, null);
		ImageIO.write(img1, "jpg", new File(dataDir + "GTOfWorksheet_out.jpg"));
		// ExEnd:1

		System.out.println("GenerateThumbnailofWorksheet executed successfully.");
	}
}
