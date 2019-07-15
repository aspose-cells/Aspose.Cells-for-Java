package AsposeCellsExamples.Files.Handling;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class OpeningFODSFiles {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the source directory.
		String sourceDir = Utils.Get_SourceDirectory();

		// Instantiate LoadOptions specified by the LoadFormat.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.FODS);

        // Create a Workbook object and opening the file from its path
        Workbook workbook = new Workbook(sourceDir + "SampleFods.fods", loadOptions);

		// Print message
		System.out.println("FODS file opened successfully!");
		// ExEnd:1
	}
}
