/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.cells.examples.data.addon.hyperlinks;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class AddingLinkToExternalFile {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AddingLinkToExternalFile.class);

        //Instantiating a Workbook object
        Workbook workbook = new Workbook();

        //Obtaining the reference of the first worksheet.
        WorksheetCollection worksheets = workbook.getWorksheets();
        Worksheet sheet = worksheets.get(0);

        //Setting a value to the "A1" cell
        Cells cells = sheet.getCells();
        Cell cell = cells.get("A1");
        cell.setValue("Visit Aspose");

        //Setting the font color of the cell to Blue
        Style style = cell.getStyle();
        style.getFont().setColor(Color.getBlue());

        //Setting the font of the cell to Single Underline
        style.getFont().setUnderline(FontUnderlineType.SINGLE);
        cell.setStyle(style);

        HyperlinkCollection hyperlinks = sheet.getHyperlinks();

        //Adding a link to the external file
        hyperlinks.add("A5", 1, 1, dataDir + "book1.xls");

        //Saving the Excel file
        workbook.save(dataDir + "output.xls");

        // Print message
        System.out.println("Process completed successfully");

    }
}
