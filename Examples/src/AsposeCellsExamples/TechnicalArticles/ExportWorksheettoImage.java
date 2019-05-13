package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.ImageType;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class ExportWorksheettoImage {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ExportWorksheettoImage.class) + "TechnicalArticles/";
		// Create workbook object from source file
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		/*
		 * Set image or print options, We want one page per sheet, The image format is in png And desired dimensions are
		 * 400x400
		 */
		ImageOrPrintOptions opts = new ImageOrPrintOptions();
		opts.setOnePagePerSheet(true);
		opts.setImageType(ImageType.PNG);
		opts.setDesiredSize(400, 400);

		// Render sheet into image
		SheetRender sr = new SheetRender(worksheet, opts);
		sr.toImage(0, dataDir + "EWSheetToImage_out.png");
        // ExEnd:1

		System.out.println("ExportWorksheettoImage executed successfully.");
	}
}
