package AsposeCellsExamples.Data;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ImportingFromJson {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ImportingFromJson.class) + "Data/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();
		Worksheet worksheet = workbook.getWorksheets().get(0);
		
		// Read File
		File file = new File(dataDir + "Test.json");
		BufferedReader bufferedReader = new BufferedReader(new FileReader(file));
        String jsonInput = "";
        String tempString;
        while ((tempString = bufferedReader.readLine()) != null) {
        	jsonInput = jsonInput + tempString; 
        }
        bufferedReader.close();
		
        // Set Styles
        CellsFactory factory = new CellsFactory();
        Style style = factory.createStyle();
        style.setHorizontalAlignment(TextAlignmentType.CENTER);
        style.getFont().setColor(Color.getBlueViolet());
        style.getFont().setBold(true);
		
        // Set JsonLayoutOptions
        JsonLayoutOptions options = new JsonLayoutOptions();
        options.setTitleStyle(style);
        options.setArrayAsTable(true);

        // Import JSON Data
        JSONUtility.importData(jsonInput, worksheet.getCells(), 0, 0, options);

        // Save Excel file
        workbook.save(dataDir + "ImportingFromJson.out.xlsx");
        // ExEnd:1
        
        System.out.println("ImportingFromJson executed successfully.");
	}
}
