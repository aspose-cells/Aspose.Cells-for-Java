package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.ImageType;
import com.aspose.cells.PrintingPageType;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class RemoveWhitespaceAroundData {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(RemoveWhitespaceAroundData.class) + "TechnicalArticles/";

		// Instantiate a workbook
		// Open the template file
		Workbook book = new Workbook(dataDir + "book1.xlsx");

		// Get the first worksheet
		Worksheet sheet = book.getWorksheets().get(0);

		// Specify your print area if you want
		// sheet.PageSetup.PrintArea = "A1:H8";

		// To remove the white border around the image.
		sheet.getPageSetup().setLeftMargin(0);
		sheet.getPageSetup().setRightMargin(0);
		sheet.getPageSetup().setTopMargin(0);
		sheet.getPageSetup().setBottomMargin(0);

		// Define ImageOrPrintOptions
		ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
		imgOptions.setImageType(ImageType.EMF);
		// Set only one page would be rendered for the image
		imgOptions.setOnePagePerSheet(true);
		imgOptions.setPrintingPage(PrintingPageType.IGNORE_BLANK);

		// Create the SheetRender object based on the sheet with its
		// ImageOrPrintOptions attributes
		SheetRender render = new SheetRender(sheet, imgOptions);
		// Convert the image
		render.toImage(0, dataDir + "RWhitespaceAroundData_out.emf");
        // ExEnd:1

		System.out.println("RemoveWhitespaceAroundData executed successfully.");
	}
}
