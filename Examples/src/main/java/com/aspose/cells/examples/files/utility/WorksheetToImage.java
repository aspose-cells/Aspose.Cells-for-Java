/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.cells.examples.files.utility;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class WorksheetToImage {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(WorksheetToImage.class);

        //Instantiate a new workbook with path to an Excel file
        Workbook book = new Workbook(dataDir + "MyTestBook1.xls");

        //Create an object for ImageOptions
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

        //Set the image type
        imgOptions.setImageFormat(ImageFormat.getPng());

        //Get the first worksheet.
        Worksheet sheet = book.getWorksheets().get(0);

        //Create a SheetRender object for the target sheet
        SheetRender sr = new SheetRender(sheet, imgOptions);
        for (int j = 0; j < sr.getPageCount(); j++) {
            //Generate an image for the worksheet
            sr.toImage(j, dataDir + "mysheetimg_" + j + ".png");

        }

        // Print message
        System.out.println("Images generated successfully.");
    }
}
