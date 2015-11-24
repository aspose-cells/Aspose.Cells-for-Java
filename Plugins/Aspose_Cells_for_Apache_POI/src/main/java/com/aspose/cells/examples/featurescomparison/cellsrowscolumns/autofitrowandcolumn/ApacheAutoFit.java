package com.aspose.cells.examples.featurescomparison.cellsrowscolumns.autofitrowandcolumn;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.aspose.cells.examples.Utils;

public class ApacheAutoFit
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ApacheAutoFit.class);

        InputStream inStream = new FileInputStream(dataDir + "workbook.xls");
        Workbook workbook = WorkbookFactory.create(inStream);

        Sheet sheet = workbook.createSheet("new sheet");
        sheet.autoSizeColumn(0); //adjust width of the first column
        sheet.autoSizeColumn(1); //adjust width of the second column

        FileOutputStream fileOut;
        fileOut = new FileOutputStream(dataDir + "AutoFit_Apache_Out.xls");
        workbook.write(fileOut);
        fileOut.close();

        System.out.println("Process Completed.");
    }
}