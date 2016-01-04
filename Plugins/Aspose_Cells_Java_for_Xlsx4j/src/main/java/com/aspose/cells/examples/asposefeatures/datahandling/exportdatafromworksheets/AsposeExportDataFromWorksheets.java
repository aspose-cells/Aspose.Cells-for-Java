package com.aspose.cells.examples.asposefeatures.datahandling.exportdatafromworksheets;

import java.io.FileInputStream;
import java.util.Arrays;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AsposeExportDataFromWorksheets
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeExportDataFromWorksheets.class);

        //Creating a file stream containing the Excel file to be opened
        FileInputStream fstream = new FileInputStream(dataDir + "workbook.xls");

        //Instantiating a Workbook object
        Workbook workbook = new Workbook(fstream);

        //Accessing the first worksheet in the Excel file
        Worksheet worksheet = workbook.getWorksheets().get(0);

        //Exporting the contents of 7 rows and 2 columns starting from 1st cell to Array.
        Object dataTable [][] = worksheet.getCells().exportArray(4,0,7,8);

        for (int i = 0 ; i < dataTable.length ; i++)
        {
                System.out.println("["+ i +"]: "+ Arrays.toString(dataTable[i]));
        }
        //Closing the file stream to free all resources
        fstream.close();
    }
}
