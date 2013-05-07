/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithworksheets.displayfeatures.splitpanes.java;

import com.aspose.cells.*;

public class SplitPanes
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithworksheets/displayfeatures/splitpanes/data/";
        
        //Instantiate a new workbook
        //Open a template file
        Workbook book = new Workbook(dataDir + "book.xls");

        //Set the active cell
        book.getWorksheets().get(0).setActiveCell("A20");

        //Split the worksheet window
        book.getWorksheets().get(0).split();

        //Save the excel file
        book.save(dataDir + "book.out.xls", SaveFormat.EXCEL_97_TO_2003);
        
        //Print Message
        System.out.println("Panes split successfully.");
    }
}




