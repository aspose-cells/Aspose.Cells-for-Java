package AsposeCellsExamples.Files.Handling;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class OpeningSXCFiles {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the source directory.
		String sourceDir = Utils.Get_SourceDirectory();

		// Instantiate LoadOptions specified by the LoadFormat.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.SXC);

        // Create a Workbook object and opening the file from its path
        Workbook workbook = new Workbook(sourceDir + "SampleSXC.sxc", loadOptions);
        
        // Using the Sheet 1 in Workbook
        Worksheet worksheet = workbook.getWorksheets().get(0);

        // Accessing a cell using its name
        Cell cell = worksheet.getCells().get("C3");

        System.out.println("Cell Name: " + cell.getName() + " Value: " + cell.getStringValue());
		// ExEnd:1
        
		System.out.println("OpeningSXCFiles executed successfully!");
	}
}
