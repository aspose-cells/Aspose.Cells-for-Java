package com.aspose.cells.examples.asposefeatures.worksheets.displayandhidescrollbars;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class AsposeDisplayAndHideScrollbars
{
    public static void main(String[] args) throws Exception
    {
	// The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeDisplayAndHideScrollbars.class);
	
	//Instantiating a Excel object by excel file path
	Workbook workbook = new Workbook(dataDir + "book1.xls");

	//Hiding the vertical scroll bar of the Excel file
	workbook.getSettings().setVScrollBarVisible(false);

	//Hiding the horizontal scroll bar of the Excel file
	workbook.getSettings().setHScrollBarVisible(false);

	//Saving the modified Excel file in default (that is Excel 2003) format
	workbook.save(dataDir + "SrollbarsHide.xls");

	// ===============================================================
	
	//Displaying the vertical scroll bar of the Excel file
	workbook.getSettings().setVScrollBarVisible(true);

	//Displaying the horizontal scroll bar of the Excel file
	workbook.getSettings().setHScrollBarVisible(true);

	//Saving the modified Excel file in default (that is Excel 2003) format
	workbook.save(dataDir + "DisplaySrollbars.xls");
	
	System.out.println("Scrollbars. Done");
    }
}