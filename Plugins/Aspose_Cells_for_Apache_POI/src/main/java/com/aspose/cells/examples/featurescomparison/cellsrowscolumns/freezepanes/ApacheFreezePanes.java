package com.aspose.cells.examples.featurescomparison.cellsrowscolumns.freezepanes;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.aspose.cells.examples.Utils;

public class ApacheFreezePanes
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ApacheFreezePanes.class);

        Workbook wb = new HSSFWorkbook();
        Sheet sheet1 = wb.createSheet("new sheet");
        Sheet sheet2 = wb.createSheet("second sheet");
        Sheet sheet3 = wb.createSheet("third sheet");

        // Freeze just one row
        sheet1.createFreezePane( 0, 2, 0, 2 );
        // Freeze just one column
        sheet2.createFreezePane( 2, 0, 2, 0 );
        // Freeze the columns and rows (forget about scrolling position of the lower right quadrant).
        sheet3.createFreezePane( 2, 2 );

        FileOutputStream fileOut = new FileOutputStream(dataDir + "ApacheFreezePanes_Out.xls");
        wb.write(fileOut);
        fileOut.close();

        //Print Message
        System.out.println("Panes freeze successfull.");
    }
}
