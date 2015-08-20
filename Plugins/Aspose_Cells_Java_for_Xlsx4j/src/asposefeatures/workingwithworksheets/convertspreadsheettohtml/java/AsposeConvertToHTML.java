package asposefeatures.workingwithworksheets.convertspreadsheettohtml.java;

import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;

public class AsposeConvertToHTML
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/asposefeatures/workingwithworksheets/convertspreadsheettohtml/data/";

		// Specify the HTML Saving Options
		HtmlSaveOptions save = new HtmlSaveOptions(SaveFormat.HTML);

		// Instantiate a workbook and open the template XLSX file
		Workbook book = new Workbook(dataPath + "book1.xls");

		// Save the HTML file
		book.save(dataPath + "AsposeHTMLSpreadsheet.html", save);

		System.out.println("Spreadsheet->HTML. Done.");
	}
}