package com.aspose.cells.examples.featurescomparison.cellsrowscolumns.iterate;

import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.aspose.cells.examples.Utils;

public class ApacheIterateRowsnCols
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ApacheIterateRowsnCols.class);

        InputStream inStream = new FileInputStream(dataDir + "workbook.xls");
        Workbook wb = WorkbookFactory.create(inStream);
        Sheet sheet = wb.getSheetAt(0);
        for (Row row : sheet) 
        {
          for (Cell cell : row) 
          {
            System.out.println("Iteration.");
          }
        }
    }
}