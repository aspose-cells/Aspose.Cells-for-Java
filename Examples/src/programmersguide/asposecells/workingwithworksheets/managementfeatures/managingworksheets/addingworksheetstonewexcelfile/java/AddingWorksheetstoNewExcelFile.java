/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithworksheets.managementfeatures.managingworksheets.addingworksheetstonewexcelfile.java;

import com.aspose.cells.*;

public class AddingWorksheetstoNewExcelFile
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithworksheets/managementfeatures/managingworksheets/addingworksheetstonewexcelfile/data/";
        
        //Instantiating a Workbook object
        Workbook workbook = new Workbook();

        //Adding a new worksheet to the Workbook object
        WorksheetCollection worksheets = workbook.getWorksheets();

        int sheetIndex = worksheets.add();
        Worksheet worksheet = worksheets.get(sheetIndex);

        //Setting the name of the newly added worksheet
        worksheet.setName("My Worksheet");

        //Saving the Excel file
        workbook.save(dataDir  + "book.out.xls");
        
        //Print Message
        System.out.println("Sheet added successfully.");
    }
}