package com.aspose.cells.examples.featurescomparison.worksheets.reordersheets;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import com.aspose.cells.examples.Utils;

public class ApacheMoveSheet
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ApacheMoveSheet.class);

        Workbook wb = new HSSFWorkbook();
        wb.createSheet("new sheet");
        wb.createSheet("second sheet");
        wb.createSheet("third sheet");

        wb.setSheetOrder("second sheet", 0);
        wb.setSheetOrder("new sheet", 1);
        wb.setSheetOrder("third sheet", 2);

        FileOutputStream fileOut = new FileOutputStream(dataDir + "ApacheMoveSheet.xls");
        wb.write(fileOut);
        fileOut.close();

        //Print Message
        System.out.println("Reordered successfull.");
    }
}
