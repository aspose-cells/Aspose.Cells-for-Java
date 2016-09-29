package com.aspose.cells.examples.articles;

import com.aspose.cells.HTMLLoadOptions;
import com.aspose.cells.LoadFormat;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class KeepPrecisionOfLargeNumbers {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(KeepPrecisionOfLargeNumbers.class) + "articles/";

		// Sample Html containing large number with digits greater than 15
		String html = "<html>" + "<body>" + "<p>1234567890123456</p>" + "</body>" + "</html>";

		// Convert Html to byte array
		byte[] byteArray = html.getBytes();

		// Set Html load options and keep precision true
		HTMLLoadOptions loadOptions = new HTMLLoadOptions(LoadFormat.HTML);
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

	}

}
