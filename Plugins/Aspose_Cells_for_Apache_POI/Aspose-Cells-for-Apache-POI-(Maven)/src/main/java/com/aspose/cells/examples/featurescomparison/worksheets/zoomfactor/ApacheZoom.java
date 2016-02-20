package com.aspose.cells.examples.featurescomparison.worksheets.zoomfactor;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.aspose.cells.examples.Utils;

public class ApacheZoom
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ApacheZoom.class);

        Workbook wb = new HSSFWorkbook();
        Sheet sheet1 = wb.createSheet("new sheet");
        sheet1.setZoom(3,4);   // 75 percent magnification

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream(dataDir + "ApacheZoom_Out.xls");
        wb.write(fileOut);
        fileOut.close();	

        System.out.println("Process Completed Successfully.");
    }
}
