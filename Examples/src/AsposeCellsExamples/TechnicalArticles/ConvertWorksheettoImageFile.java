package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.ImageType;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class ConvertWorksheettoImageFile {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ConvertWorksheettoImageFile.class) + "TechnicalArticles/";
		// Create a new Workbook object
		// Open a template excel file
		Workbook book = new Workbook(dataDir + "book1.xlsx");
		// Get the first worksheet
		Worksheet sheet = book.getWorksheets().get(0);

		// Define ImageOrPrintOptions
		ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
		// Specify the image format
		imgOptions.setImageType(ImageType.JPEG);

		// Render the sheet with respect to specified image/print options
		SheetRender render = new SheetRender(sheet, imgOptions);
		// Render the image for the sheet
		render.toImage(0, dataDir + "CWToImageFile.jpg");
        // ExEnd:1

		System.out.println("ConvertWorksheettoImageFile executed successfully.");
	}
}
