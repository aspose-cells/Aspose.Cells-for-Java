package com.aspose.cells.examples.featurescomparison.workbook.createnewworkbook;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import com.aspose.cells.examples.Utils;

public class ApacheNewWorkbook
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ApacheNewWorkbook.class);

        Workbook wb = new HSSFWorkbook();

        FileOutputStream fileOut;
        fileOut = new FileOutputStream(dataDir + "ApacheNewWorkBook.xls");
        wb.write(fileOut);
        fileOut.close();

        System.out.println("File Created.");
    }
}