package com.aspose.cells.examples.featurescomparison.cellsrowscolumns.splitpanes;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

import com.aspose.cells.examples.Utils;

/**
* How to set split panes
*/
public class ApacheSplitPanes 
{
    public static void main(String[]args) throws Exception 
    { 
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ApacheSplitPanes.class);

        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");

        // Create a split with the lower left side being the active quadrant
        sheet.createSplitPane(2000, 2000, 0, 0, Sheet.PANE_LOWER_LEFT);

        FileOutputStream fileOut = new FileOutputStream(dataDir + "ApacheSplitFreezePanes.xlsx");
        wb.write(fileOut);
        fileOut.close();

        System.out.println("Done.");
    }
}