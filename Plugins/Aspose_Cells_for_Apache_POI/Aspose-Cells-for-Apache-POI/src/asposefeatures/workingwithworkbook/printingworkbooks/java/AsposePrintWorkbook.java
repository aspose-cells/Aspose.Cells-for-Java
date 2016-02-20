package asposefeatures.workingwithworkbook.printingworkbooks.java;

import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.WorkbookRender;
import com.aspose.cells.Worksheet;

public class AsposePrintWorkbook
{
    public static void main(String[] args) throws Exception
    {
	String dataPath = "src/asposefeatures/workingwithworkbook/printingworkbooks/data/";

	// Instantiate a new workbook
	Workbook book = new Workbook(dataPath + "AsposeDataInput.xls");

	// Create an object for ImageOptions
	ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

	// Get the first worksheet
	Worksheet sheet = book.getWorksheets().get(0);

	// Create a SheetRender object with respect to your desired sheet
	SheetRender sr = new SheetRender(sheet, imgOptions);

	// Print the worksheet
	sr.toPrinter("Samsung ML-1520 Series");

	// Create a WorkbookRender object with respect to your workbook
	WorkbookRender wr = new WorkbookRender(book, imgOptions);

	// Print the workbook
	wr.toPrinter("Samsung ML-1520 Series");
    }
}
