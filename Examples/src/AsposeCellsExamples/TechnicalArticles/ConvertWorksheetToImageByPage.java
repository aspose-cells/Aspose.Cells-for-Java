package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.ImageType;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class ConvertWorksheetToImageByPage {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ConvertWorksheetToImageByPage.class) + "TechnicalArticles/";
		// Create a new Workbook object
		// Open a template excel file
		Workbook book = new Workbook(dataDir + "ConvertWorksheetToImageByPage.xlsx");
		// Get the first worksheet
		Worksheet sheet = book.getWorksheets().get(0);
		// Define ImageOrPrintOptions
		ImageOrPrintOptions options = new ImageOrPrintOptions();
		// Set Resolution
		options.setHorizontalResolution(200);
		options.setVerticalResolution(200);
		options.setImageType(ImageType.TIFF);

		// Sheet2Image by page conversion
		SheetRender render = new SheetRender(sheet, options);
		for (int j = 0; j < render.getPageCount(); j++) {
			render.toImage(j, dataDir + sheet.getName() + " Page" + (j + 1) + ".tif");
		}
        // ExEnd:1

		System.out.println("ConvertWorksheetToImageByPage executed successfully.");
	}
}
