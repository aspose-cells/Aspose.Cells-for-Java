package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

public class ConvertCsvToJson {
    public static void main(String[] args) throws Exception {
        // ExStart:1
        //Source directory
        String sourceDir = Utils.Get_SourceDirectory();

        LoadOptions loadOptions = new LoadOptions(LoadFormat.CSV);
        // Load CSV file
        Workbook workbook = new Workbook(sourceDir + "SampleCsv.csv", loadOptions);
        Cell lastCell = workbook.getWorksheets().get(0).getCells().getLastCell();

        // Set ExportRangeToJsonOptions
        ExportRangeToJsonOptions options = new ExportRangeToJsonOptions();
        Range range = workbook.getWorksheets().get(0).getCells().createRange(0, 0, lastCell.getRow() + 1, lastCell.getColumn() + 1);
        String data = JsonUtility.exportRangeToJson(range, options);

        // Print JSON
        System.out.println(data);
        // ExEnd:1

        System.out.println("ConvertCsvToJson executed successfully.");
    }
}
