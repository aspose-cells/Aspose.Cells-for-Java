/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithdata.addonfeatures.namedranges.accessspecificnamedrange.java;

import com.aspose.cells.*;

public class AccessSpecificNamedRange
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithdata/addonfeatures/namedranges/accessspecificnamedrange/data/";

        //Instantiating a Workbook object
        Workbook workbook = new Workbook(dataDir + "book1.xls");

        WorksheetCollection worksheets = workbook.getWorksheets();

        //Accessing the first worksheet in the Excel file
        Worksheet sheet = worksheets.get(0);
        Cells cells = sheet.getCells();

        //Getting the specified named range
        Range namedRange = worksheets.getRangeByName("TestRange");

        // Print message
        System.out.println("Named Range : " + namedRange.getRefersTo());
    }
}




