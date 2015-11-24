package com.aspose.cells.examples.featurescomparison.workbook.fittoonepage;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.aspose.cells.examples.Utils;

public class ApacheFitSheetToOnePage {
    public static void main(String[]args) throws Exception 
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ApacheFitSheetToOnePage.class);

        Workbook wb = new XSSFWorkbook();  //or new HSSFWorkbook();
        Sheet sheet = wb.createSheet("format sheet");
        PrintSetup ps = sheet.getPrintSetup();

        sheet.setAutobreaks(true);

        ps.setFitHeight((short) 1);
        ps.setFitWidth((short) 1);

        // Create various cells and rows for spreadsheet.

        FileOutputStream fileOut = new FileOutputStream(dataDir + "ApacheFitSheet.xlsx");
        wb.write(fileOut);
        fileOut.close();

        System.out.println("Done.");
    }
}