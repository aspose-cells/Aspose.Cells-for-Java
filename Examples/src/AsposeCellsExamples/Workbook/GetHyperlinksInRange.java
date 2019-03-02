package AsposeCellsExamples.Workbook;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class GetHyperlinksInRange {
	
	static String sourceDir = Utils.Get_SourceDirectory();
	static String outputDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		// ExStart:1
        // Instantiate a Workbook object
        // Open an Excel file
        Workbook workbook = new Workbook(sourceDir + "HyperlinksSample.xlsx");

        // Get the first (default) worksheet
        Worksheet worksheet = workbook.getWorksheets().get(0);

        // Create a range A2:B3
        Range range = worksheet.getCells().createRange("A2", "B3");

        // Get Hyperlinks in range
        Hyperlink[] hyperlinks = range.getHyperlinks();

        for (Hyperlink link : hyperlinks){
            System.out.println(link.getArea() + " : " + link.getAddress());

            // To delete the link, use the Hyperlink.Delete() method.
            link.delete();
        }

        workbook.save(outputDir + "HyperlinksSample_out.xlsx");
        // ExEnd:1
	}
}
