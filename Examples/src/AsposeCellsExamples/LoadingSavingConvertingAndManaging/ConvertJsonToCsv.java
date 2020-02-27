package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

import java.nio.file.Files;
import java.nio.file.Paths;

public class ConvertJsonToCsv {
    public static void main(String[] args) throws Exception {
        // ExStart:1
        //Source directory
        String sourceDir = Utils.Get_SourceDirectory();

        //Output directory
        String outputDir = Utils.Get_OutputDirectory();

        // Read JSON file
        String str = new String(Files.readAllBytes(Paths.get(sourceDir + "SampleJson.json")));

        // Create empty workbook
        Workbook workbook = new Workbook();

        // Get Cells
        Cells cells = workbook.getWorksheets().get(0).getCells();

        // Set JsonLayoutOptions
        JsonLayoutOptions importOptions = new JsonLayoutOptions();
        importOptions.setConvertNumericOrDate(true);
        importOptions.setArrayAsTable(true);
        importOptions.setIgnoreArrayTitle(true);
        importOptions.setIgnoreObjectTitle(true);
        JsonUtility.importData(str, cells, 0, 0, importOptions);

        // Save Workbook
        workbook.save(outputDir + "SampleJson_out.csv");
        // ExEnd:1

        System.out.println("ConvertJsonToCsv executed successfully.");
    }
}
