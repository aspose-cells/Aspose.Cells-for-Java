/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.cells.examples.worksheets.security.unprotecting;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class UnprotectingSimplyProtectedWorksheet {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(UnprotectingSimplyProtectedWorksheet.class);

        //Instantiating a Workbook object
        Workbook workbook = new Workbook(dataDir + "book1.xls");

        WorksheetCollection worksheets = workbook.getWorksheets();
        Worksheet worksheet = worksheets.get(0);

        Protection protection = worksheet.getProtection();

        //The following 3 methods are only for Excel 2000 and earlier formats
        protection.setAllowEditingContent(false);
        protection.setAllowEditingObject(false);
        protection.setAllowEditingScenario(false);

        //Unprotecting the worksheet
        worksheet.unprotect();

        // Save the excel file.
        workbook.save(dataDir + "output.xls", FileFormatType.EXCEL_97_TO_2003);

        //Print Message
        System.out.println("Worksheet unprotected successfully.");
    }
}
