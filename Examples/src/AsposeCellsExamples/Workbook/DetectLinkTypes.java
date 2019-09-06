package AsposeCellsExamples.Workbook;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

public class DetectLinkTypes {

    // ExStart:1
	public static void main(String[] args) throws Exception {
        // The path to the directories.
        String sourceDir = Utils.Get_SourceDirectory();

        Workbook workbook = new Workbook(sourceDir + "LinkTypes.xlsx");

        // Get the first (default) worksheet
        Worksheet worksheet = workbook.getWorksheets().get(0);

        // Create a range A2:B3
        Range range = worksheet.getCells().createRange("A1", "A7");

        // Get Hyperlinks in range
        Hyperlink[] hyperlinks = range.getHyperlinks();

        for (Hyperlink link : hyperlinks)
        {
            System.out.println(link.getTextToDisplay() + ": " + getLinkTypeName(link.getLinkType()));
        }

		System.out.println("DetectLinkTypes executed successfully.");
	}

	private static String getLinkTypeName(int linkType){
	    if(linkType == TargetModeType.EXTERNAL){
	        return "EXTERNAL";
        } else if(linkType == TargetModeType.FILE_PATH){
	        return "FILE_PATH";
        } else if(linkType == TargetModeType.EMAIL){
            return "EMAIL";
        } else {
	        return "CELL_REFERENCE";
        }
    }
    // ExEnd:1
}
