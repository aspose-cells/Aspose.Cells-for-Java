package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.ImageType;
import com.aspose.cells.Picture;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class ExtractImagesfromWorksheets {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ExtractImagesfromWorksheets.class) + "TechnicalArticles/";
		// Open a template Excel file
		Workbook workbook = new Workbook(dataDir + "book3.xlsx");

		// Get the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Get the first Picture in the first worksheet
		Picture pic = worksheet.getPictures().get(0);

		// Set the output image file path
		String fileName = "aspose-logo.jpg";

		// Note: you may evaluate the image format before specifying the image path

		// Define ImageOrPrintOptions
		ImageOrPrintOptions printoption = new ImageOrPrintOptions();

		// Specify the image format
		printoption.setImageType(ImageType.JPEG);

		// Save the image
		pic.toImage(dataDir + fileName, printoption);
        // ExEnd:1

		System.out.println("ExtractImagesfromWorksheets executed successfully.");
	}
}
