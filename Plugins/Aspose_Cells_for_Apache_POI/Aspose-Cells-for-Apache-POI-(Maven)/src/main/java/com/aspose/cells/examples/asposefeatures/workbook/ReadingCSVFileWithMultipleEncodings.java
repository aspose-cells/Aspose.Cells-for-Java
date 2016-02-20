package com.aspose.cells.examples.asposefeatures.workbook;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.TxtLoadOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ReadingCSVFileWithMultipleEncodings
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ReadingCSVFileWithMultipleEncodings.class);

        //Set Multi Encoded Property to True
        TxtLoadOptions options = new TxtLoadOptions();
        options.setMultiEncoded(true);

        //Load the CSV file into Workbook
        Workbook workbook = new Workbook(dataDir + "MultiEncoded.csv", options);

        //Save it in XLSX format
        workbook.save(dataDir + "EncodedNewFile_Out.xlsx", SaveFormat.XLSX);

        System.out.println("MultiEncoded file successfully read.");
    }
}
