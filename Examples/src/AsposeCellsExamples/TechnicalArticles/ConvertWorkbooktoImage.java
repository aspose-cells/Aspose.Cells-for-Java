package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.ImageType;
import com.aspose.cells.Workbook;
import com.aspose.cells.WorkbookRender;
import AsposeCellsExamples.Utils;

public class ConvertWorkbooktoImage {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ConvertWorkbooktoImage.class) + "TechnicalArticles/";
		// Instantiate a new Workbook object
		Workbook book = new Workbook(dataDir + "book1.xlsx");

		// Apply different Image and Print options
		ImageOrPrintOptions options = new ImageOrPrintOptions();

		// Set Image Format
		options.setImageType(ImageType.TIFF);

		// If you want entire sheet as a single image
		options.setOnePagePerSheet(true);

		// Render to image
		WorkbookRender render = new WorkbookRender(book, options);
		render.toImage(dataDir + "CWorkbooktoImage_out.tiff");
        // ExEnd:1

		System.out.println("ConvertWorkbooktoImage executed successfully.");
	}
}
