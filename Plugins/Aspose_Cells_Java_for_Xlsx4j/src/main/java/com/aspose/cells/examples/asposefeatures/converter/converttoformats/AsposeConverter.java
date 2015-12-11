package com.aspose.cells.examples.asposefeatures.converter.converttoformats;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class AsposeConverter
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeConverter.class);

        Workbook workbook = new Workbook(dataDir + "workbook.xls");

        //Save the document in PDF format
        workbook.save(dataDir + "AsposeConvert_Out.pdf", SaveFormat.PDF);

        // Print message
        System.out.println("Excel to PDF conversion performed successfully.");
    }
}
