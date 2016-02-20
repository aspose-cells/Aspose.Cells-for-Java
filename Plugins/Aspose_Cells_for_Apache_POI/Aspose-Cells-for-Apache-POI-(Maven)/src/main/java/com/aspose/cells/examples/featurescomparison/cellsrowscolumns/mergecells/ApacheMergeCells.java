package com.aspose.cells.examples.featurescomparison.cellsrowscolumns.mergecells;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import com.aspose.cells.examples.Utils;

public class ApacheMergeCells
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ApacheMergeCells.class);

        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");

        Row row = sheet.createRow((short) 1);
        Cell cell = row.createCell((short) 1);
        cell.setCellValue("This is a test of merging");

        sheet.addMergedRegion(new CellRangeAddress(
                1, //first row (0-based)
                1, //last row  (0-based)
                1, //first column (0-based)
                2  //last column  (0-based)
        ));

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream(dataDir + "ApacheMergeCells_Out.xls");
        wb.write(fileOut);
        fileOut.close();

        // Print message
        System.out.println("Process completed successfully");
    }
}
