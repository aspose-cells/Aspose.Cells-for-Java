package com.aspose.cells.examples.featurescomparison.worksheets.reordersheets;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class AsposeMoveSheet
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeMoveSheet.class);
		
	//Create a new Workbook.
	Workbook workbook = new Workbook();

	WorksheetCollection worksheets = workbook.getWorksheets();
        Worksheet worksheet1 = worksheets.get(0);
        Worksheet worksheet2 = worksheets.add("Sheet2");
        Worksheet worksheet3 = worksheets.add("Sheet3");
        
	//Move Sheets with in Workbook.
        worksheet2.moveTo(0);
        worksheet1.moveTo(1);
        worksheet3.moveTo(2);

	//Save the excel file.
        workbook.save(dataDir + "AsposeMoveSheet.xls");
		
	System.out.println("Sheet moved successfully."); // Print Message
    }
}
