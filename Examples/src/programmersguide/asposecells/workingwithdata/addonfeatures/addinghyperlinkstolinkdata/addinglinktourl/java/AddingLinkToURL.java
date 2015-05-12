/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithdata.addonfeatures.addinghyperlinkstolinkdata.addinglinktourl.java;

import com.aspose.cells.*;

public class AddingLinkToURL
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithdata/addonfeatures/addinghyperlinkstolinkdata/addinglinktourl/data/";

        //Instantiating a Workbook object
        Workbook workbook = new Workbook();

        //Obtaining the reference of the first worksheet.
        WorksheetCollection worksheets = workbook.getWorksheets();
        Worksheet sheet = worksheets.get(0);
        HyperlinkCollection hyperlinks = sheet.getHyperlinks();

        //Adding a hyperlink to a URL at "A1" cell
        hyperlinks.add("A1",1,1,"http://www.aspose.com");

        //Saving the Excel file
        workbook.save(dataDir + "output.xls");

        // Print message
        System.out.println("Process completed successfully");

    }
}




