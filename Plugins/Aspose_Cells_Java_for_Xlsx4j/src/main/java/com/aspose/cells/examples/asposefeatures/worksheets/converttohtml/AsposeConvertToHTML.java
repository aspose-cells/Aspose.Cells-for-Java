package com.aspose.cells.examples.asposefeatures.worksheets.converttohtml;

import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class AsposeConvertToHTML
{
    public static void main(String[] args) throws Exception
    {
	// The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeConvertToHTML.class);
	
	//Specify the HTML Saving Options
	HtmlSaveOptions save = new HtmlSaveOptions(SaveFormat.HTML);

	//Instantiate a workbook and open the template XLSX file
	Workbook book = new Workbook(dataDir + "book1.xls");

	//Save the HTML file
	book.save(dataDir + "AsposeHTMLSpreadsheet.html", save);
	
	System.out.println("Spreadsheet->HTML. Done.");
    }
}