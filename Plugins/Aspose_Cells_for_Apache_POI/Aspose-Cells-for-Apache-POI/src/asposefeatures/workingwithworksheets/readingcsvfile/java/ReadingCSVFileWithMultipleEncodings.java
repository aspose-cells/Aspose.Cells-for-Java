package asposefeatures.workingwithworksheets.readingcsvfile.java;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.TxtLoadOptions;
import com.aspose.cells.Workbook;

public class ReadingCSVFileWithMultipleEncodings
{
    public static void main(String[] args) throws Exception
    {
	String dataPath = "src/asposefeatures/workingwithworksheets/readingcsvfile/data/";

	// Set Multi Encoded Property to True
	TxtLoadOptions options = new TxtLoadOptions();
	options.setMultiEncoded(true);

	// Load the CSV file into Workbook
	Workbook workbook = new Workbook(dataPath + "MultiEncoded.csv", options);

	// Save it in XLSX format
	workbook.save(dataPath + "EncodedNewFile_Out.xlsx", SaveFormat.XLSX);

	System.out.println("MultiEncoded file successfully read.");
    }
}
