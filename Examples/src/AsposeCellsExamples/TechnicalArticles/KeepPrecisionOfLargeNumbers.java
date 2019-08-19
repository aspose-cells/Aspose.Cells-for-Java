package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class KeepPrecisionOfLargeNumbers {

	public static void main(String[] args) throws Exception {

		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(KeepPrecisionOfLargeNumbers.class) + "TechnicalArticles/";

		// Sample Html containing large number with digits greater than 15
		String html = "<html>" + "<body>" + "<p>1234567890123456</p>" + "</body>" + "</html>";

		// Convert Html to byte array
		byte[] byteArray = html.getBytes();

		// Set Html load options and keep precision true
		HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.HTML);
		loadOptions.setKeepPrecision(true);

		// Convert byte array into stream
		java.io.ByteArrayInputStream stream = new java.io.ByteArrayInputStream(byteArray);

		// Create workbook from stream with Html load options
		Workbook workbook = new Workbook(stream, loadOptions);

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Auto fit the sheet columns
		worksheet.autoFitColumns();

		// Save the workbook
		workbook.save(dataDir + "KPOfLargeNumbers_out.xlsx", SaveFormat.XLSX);

		System.out.println("File saved");
		// ExEnd:1
	}

}
