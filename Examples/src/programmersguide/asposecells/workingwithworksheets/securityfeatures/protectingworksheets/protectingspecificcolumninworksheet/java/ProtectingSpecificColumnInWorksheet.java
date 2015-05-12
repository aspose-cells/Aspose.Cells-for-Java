/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithworksheets.securityfeatures.protectingworksheets.protectingspecificcolumninworksheet.java;

import com.aspose.cells.*;

public class ProtectingSpecificColumnInWorksheet
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithworksheets/securityfeatures/protectingworksheets/protectingspecificcolumninworksheet/data/";

        // Create a new workbook.
        Workbook wb = new Workbook();

        // Create a worksheet object and obtain the first sheet.
        Worksheet sheet = wb.getWorksheets().get(0);

        // Define the style object.
        Style style;

        // Define the styleflag object.
        StyleFlag flag;

        // Loop through all the columns in the worksheet and unlock them.
        for(int i = 0; i <= 255; i ++)
        {
            style = sheet.getCells().getColumns().get(i).getStyle();
            style.setLocked(false);
            flag = new StyleFlag();
            flag.setLocked(true);
            sheet.getCells().getColumns().get(i).applyStyle(style, flag);
        }

        // Get the first column style.
        style = sheet.getCells().getColumns().get(0).getStyle();

        // Lock it.
        style.setLocked(true);

        // Instantiate the flag.
        flag = new StyleFlag();

        // Set the lock setting.
        flag.setLocked(true);

        // Apply the style to the first column.
        sheet.getCells().getColumns().get(0).applyStyle(style, flag);

        // Save the excel file.
        wb.save(dataDir + "lockedcolumn.xls", FileFormatType.EXCEL_97_TO_2003);
        
        //Print Message
        System.out.println("Column protected successfully.");
    }
}




