package com.aspose.cells.examples.articles;

import com.aspose.cells.HTMLLoadOptions;
import com.aspose.cells.LoadFormat;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class DeleteRedundantSpacesFromHtml {

	public static void main(String[] args) throws Exception {


		// The path to the documents directory
		String dataDir = Utils.getSharedDataDir(DeleteRedundantSpacesFromHtml.class) + "articles/";

		// Sample Html containing redundant spaces after <br> tag
		String html = "<html>" + "<body>" + "<table>" + "<tr>" + "<td>" + "<br>    This is sample data"
				+ "<br>    This is sample data" + "<br>    This is sample data" + "</td>" + "</tr>" + "</table>"
				+ "</body>" + "</html>";

		// Convert Html to byte array
		byte[] byteArray = html.getBytes();

		// Set Html load options and keep precision true
		HTMLLoadOptions loadOptions = new HTMLLoadOptions(LoadFormat.HTML);
		loadOptions.setDeleteRedundantSpaces(true);

		// Convert byte array into stream
		java.io.ByteArrayInputStream stream = new java.io.ByteArrayInputStream(byteArray);

		// Create workbook from stream with Html load options
		Workbook workbook = new Workbook(stream, loadOptions);

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Auto fit the sheet columns
		worksheet.autoFitColumns();

		// Save the workbook
		workbook.save(dataDir + "DRSFromHtml_out-" + loadOptions.getDeleteRedundantSpaces() + ".xlsx", SaveFormat.XLSX);

		System.out.println("File saved");

	}

}
