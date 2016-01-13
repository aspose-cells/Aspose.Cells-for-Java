package com.aspose.cells.examples.files.utility;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class Excel2PDFConversion {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(Excel2PDFConversion.class);

        Workbook workbook = new Workbook(dataDir + "Book1.xls");

        //Save the document in PDF format
        workbook.save(dataDir + "OutBook1.pdf", SaveFormat.PDF);

        // Print message
        System.out.println("Excel to PDF conversion performed successfully.");
    }
}
