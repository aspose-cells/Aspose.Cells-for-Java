package com.aspose.cells.examples.featurescomparison.workbook.createnewworkbook;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class AsposeNewWorkbook
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeNewWorkbook.class);

        Workbook workbook = new Workbook(); // Creating a Workbook object

        //Workbooks can be saved in many formats
        workbook.save(dataDir + "newWorkBook_Aspose_Out.xlsx", FileFormatType.XLSX);

        System.out.println("Workbook saved successfully."); // Print Message
    }
}