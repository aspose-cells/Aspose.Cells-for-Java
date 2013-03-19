/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithfiles.utilityfeatures.excel2pdfconversion.java;

import com.aspose.cells.*;

public class Excel2PDFConversion
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithfiles/utilityfeatures/excel2pdfconversion/data/";

        Workbook workbook = new Workbook(dataDir + "Book1.xls");

        //Save the document in PDF format
        workbook.save(dataDir + "OutBook1.pdf", SaveFormat.PDF);

        // Print message....
        System.out.println("Excel to PDF conversion performed successfully.");
    }
}




