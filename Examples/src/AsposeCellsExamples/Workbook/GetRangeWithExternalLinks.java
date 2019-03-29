package AsposeCellsExamples.Workbook;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class GetRangeWithExternalLinks {

	static String sourceDir = Utils.Get_SourceDirectory();
	static String outputDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {
		// ExStart:1
        // Instantiate a Workbook object and Open an Excel file
        Workbook workbook = new Workbook(sourceDir + "SampleExternalReferences.xlsx");
        Name namedRange = workbook.getWorksheets().getNames().get("Names");
       
        // Get ReferredAreas
        ReferredArea[] referredAreas = namedRange.getReferredAreas(true);
       
        if (referredAreas != null) {
        	for (int i = 0; i < referredAreas.length; i++) {
        		ReferredArea referredArea = referredAreas[i];
        		// Print the data in Referred Area
        		System.out.println("IsExternalLink: " + referredArea.isExternalLink());
        		System.out.println("IsArea: " + referredArea.isArea());
        		System.out.println("SheetName: " + referredArea.getSheetName());
        		System.out.println("ExternalFileName: " + referredArea.getExternalFileName());
        		System.out.println("StartColumn: " + referredArea.getStartColumn());
        		System.out.println("StartRow: " + referredArea.getStartRow());
        		System.out.println("EndColumn: " + referredArea.getEndColumn());
        		System.out.println("EndRow: " + referredArea.getEndRow());
        	}
        }
        // ExEnd:1
        
        System.out.println("GetRangeWithExternalLinks executed successfully.");
	}
}
