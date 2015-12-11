package com.aspose.cells.examples.asposefeatures.worksheets.displayandhidetabs;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class AsposeDisplayAndHideTabs
{
    public static void main(String[] args) throws Exception
    {
	// The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeDisplayAndHideTabs.class);
	
	//Instantiating a Workbook object by excel file path
	Workbook workbook = new Workbook(dataDir + "book1.xls");

	//Hiding the tabs of the Excel file
	workbook.getSettings().setShowTabs(false);

	//Saving the modified Excel file in default (that is Excel 2003) format
	workbook.save(dataDir + "AsposeHideTabs.xls");

	// ===============================================================
	
	//Displaying the tabs of the Excel file
	workbook.getSettings().setShowTabs(true);

	//Saving the modified Excel file in default (that is Excel 2003) format
	workbook.save(dataDir + "AsposeDisplayTabs.xls");
	
	System.out.println("Tabs. Done");
    }
}