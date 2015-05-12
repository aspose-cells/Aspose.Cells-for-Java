/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithworksheets.managementfeatures.managingworksheets.removingworksheetsusingsheetindex.java;

import java.io.FileInputStream;

import com.aspose.cells.*;

public class RemovingWorksheetsusingSheetIndex
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithworksheets/managementfeatures/managingworksheets/removingworksheetsusingsheetindex/data/";
        
        //Creating a file stream containing the Excel file to be opened
        FileInputStream fstream=new FileInputStream(dataDir + "book.xls");

        //Instantiating a Workbook object with the stream
        Workbook workbook = new Workbook(fstream);

        //Removing a worksheet using its sheet index
        workbook.getWorksheets().removeAt(0);
        
        //Saving the Excel file
        workbook.save(dataDir + "book.out.xls");

        //Closing the file stream to free all resources
        fstream.close();

        
        //Print Message
        System.out.println("Sheet removed successfully.");
    }
}