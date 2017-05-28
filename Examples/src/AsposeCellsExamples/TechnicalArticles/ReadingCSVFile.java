package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.TxtLoadOptions;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class ReadingCSVFile {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ReadingCSVFile.class) + "TechnicalArticles/";

		// Set Multi Encoded Property to True
		TxtLoadOptions options = new TxtLoadOptions();
		options.setMultiEncoded(true);

		// Load the CSV file into Workbook
		Workbook workbook = new Workbook(dataDir + "MutliEncoded.csv", options);

		// Save it in XLSX format
		workbook.save(dataDir + "ReadingCSVFile_out.xlsx", SaveFormat.XLSX);

	}
}
