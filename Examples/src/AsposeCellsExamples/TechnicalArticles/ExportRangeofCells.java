package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ExportRangeofCells {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ExportRangeofCells.class) + "TechnicalArticles/";
		// Create workbook from source file.
		Workbook workbook = new Workbook(dataDir + "book1.xlsx");

		// Access the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Set the print area with your desired range
		worksheet.getPageSetup().setPrintArea("E8:H10");

		// Set all margins as 0
		worksheet.getPageSetup().setLeftMargin(0);
		worksheet.getPageSetup().setRightMargin(0);
		worksheet.getPageSetup().setTopMargin(0);
		worksheet.getPageSetup().setBottomMargin(0);

		// Set OnePagePerSheet option as true
		ImageOrPrintOptions options = new ImageOrPrintOptions();
		options.setOnePagePerSheet(true);
		options.setImageType(ImageType.JPEG);

		// Take the image of your worksheet
		SheetRender sr = new SheetRender(worksheet, options);
		sr.toImage(0, dataDir + "ERangeofCells_out.jpg");
        // ExEnd:1

		System.out.println("ExportRangeofCells executed successfully.");
	}
}
