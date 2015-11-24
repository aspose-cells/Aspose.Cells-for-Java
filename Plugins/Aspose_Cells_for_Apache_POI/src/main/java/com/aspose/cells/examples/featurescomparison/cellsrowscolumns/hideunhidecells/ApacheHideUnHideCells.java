package com.aspose.cells.examples.featurescomparison.cellsrowscolumns.hideunhidecells;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.aspose.cells.examples.Utils;

public class ApacheHideUnHideCells
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ApacheHideUnHideCells.class);

        InputStream inStream = new FileInputStream(dataDir + "workbook.xls");
        Workbook workbook = WorkbookFactory.create(inStream);
        Sheet sheet = workbook.createSheet();
        Row row = sheet.createRow(0);
        row.setZeroHeight(true);

        FileOutputStream fileOut = new FileOutputStream(dataDir + "ApacheHideUnhideCells.xls");
        workbook.write(fileOut);
        fileOut.close();

        System.out.println("Process Completed.");
    }
}
