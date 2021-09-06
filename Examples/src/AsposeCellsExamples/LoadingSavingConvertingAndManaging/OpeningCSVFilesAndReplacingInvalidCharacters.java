package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class OpeningCSVFilesAndReplacingInvalidCharacters {

	public static void main(String[] args) throws Exception {

		// ExStart:1
        //Source directory
		String dataDir = Utils.getSharedDataDir(OpeningCSVFilesAndReplacingInvalidCharacters.class) + "LoadingSavingConvertingAndManaging/";
        
		LoadOptions loadOptions = new LoadOptions(LoadFormat.CSV);

        //Load CSV file
		Workbook workbook = new Workbook(dataDir + "[20180220142533][ASPOSE_CELLS_TEST].csv", loadOptions);

        System.out.println(workbook.getWorksheets().get(0).getName()); // (20180220142533)(ASPOSE_CELLS_T
        System.out.println(workbook.getWorksheets().get(0).getName().length()); // 31
        System.out.println("CSV file opened successfully!");
        // ExEnd:1

	}
}
