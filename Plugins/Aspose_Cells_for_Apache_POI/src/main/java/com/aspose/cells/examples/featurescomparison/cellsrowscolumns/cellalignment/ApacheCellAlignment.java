package com.aspose.cells.examples.featurescomparison.cellsrowscolumns.cellalignment;

import org.apache.poi.hssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

import com.aspose.cells.examples.Utils;

/**
* Shows how various alignment options work.
*
* @author Glen Stampoultzis (glens at apache.org)
*/
public class ApacheCellAlignment 
{
    public static void main(String[] args) throws IOException 
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ApacheCellAlignment.class);

        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet("new sheet");
        HSSFRow row = sheet.createRow(2);
        createCell(wb, row, 0, HSSFCellStyle.ALIGN_CENTER);
        createCell(wb, row, 1, HSSFCellStyle.ALIGN_CENTER_SELECTION);
        createCell(wb, row, 2, HSSFCellStyle.ALIGN_FILL);
        createCell(wb, row, 3, HSSFCellStyle.ALIGN_GENERAL);
        createCell(wb, row, 4, HSSFCellStyle.ALIGN_JUSTIFY);
        createCell(wb, row, 5, HSSFCellStyle.ALIGN_LEFT);
        createCell(wb, row, 6, HSSFCellStyle.ALIGN_RIGHT);

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream(dataDir + "ApahceAlignment.xls");
        wb.write(fileOut);
        fileOut.close();

        System.out.println("Done.");
    }

    /**
     * Creates a cell and aligns it a certain way.
     *
     * @param wb        the workbook
     * @param row       the row to create the cell in
     * @param column    the column number to create the cell in
     * @param align     the alignment for the cell.
     */
    private static void createCell(HSSFWorkbook wb, HSSFRow row, int column, int align) {
        HSSFCell cell = row.createCell(column);
        cell.setCellValue("Align It");
        HSSFCellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment((short)align);
        cell.setCellStyle(cellStyle);
    }
}