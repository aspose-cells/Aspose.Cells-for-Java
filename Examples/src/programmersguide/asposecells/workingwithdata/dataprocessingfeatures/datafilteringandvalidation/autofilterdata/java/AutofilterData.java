/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithdata.dataprocessingfeatures.datafilteringandvalidation.autofilterdata.java;

import com.aspose.cells.*;

public class AutofilterData
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithdata/dataprocessingfeatures/datafilteringandvalidation/autofilterdata/data/";

        //Instantiating a Workbook object
        Workbook workbook = new Workbook(dataDir + "book1.xls");

        //Accessing the first worksheet in the Excel file
        Worksheet worksheet = workbook.getWorksheets().get(0);

        //Creating AutoFilter by giving the cells range
        AutoFilter autoFilter = worksheet.getAutoFilter();
        CellArea area = new CellArea();
        autoFilter.setRange("A1:B1");

        //Saving the modified Excel file
        workbook.save(dataDir + "output.xls");

        // Print message
        System.out.println("Process completed successfully");
    }
}




