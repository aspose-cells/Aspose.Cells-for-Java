/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithfiles.utilityfeatures.convertingtomhtmlfiles.java;

import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;

public class ConvertingToMHTMLFiles
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithfiles/utilityfeatures/convertingtomhtmlfiles/data/";

        //Specify the file path
        String filePath = dataDir + "Book1.xlsx";

        //Specify the HTML saving options
        HtmlSaveOptions sv = new HtmlSaveOptions(SaveFormat.M_HTML);

        //Instantiate a workbook and open the template XLSX file
        Workbook wb = new Workbook(filePath);

        //Save the MHT file
        wb.save(filePath + ".out.mht", sv);

        // Print message
        System.out.println("Excel to MHTML conversion performed successfully.");
    }
}




