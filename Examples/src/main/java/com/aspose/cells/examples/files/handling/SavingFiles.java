/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.cells.examples.files.handling;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class SavingFiles {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SavingFiles.class);

        //Creating an Workbook object with an Excel file path
        Workbook workbook = new Workbook(dataDir + "book.xls");

        //Save in default (Excel2003) format
        workbook.save(dataDir + "book.default.out.xls");

        //Save in Excel2003 format
        workbook.save(dataDir + "book.out.xls", FileFormatType.EXCEL_97_TO_2003);

        //Save in Excel2007 xlsx format
        workbook.save(dataDir + "book.out.xlsx", FileFormatType.XLSX);

        //Save in SpreadsheetML format
        workbook.save(dataDir + "book.out.xml", FileFormatType.EXCEL_2003_XML);

        //Print Message
        System.out.println("Worksheets are saved successfully.");
    }
}
