package com.aspose.cells.examples.data.addon.namedranges;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

import java.util.ArrayList;

public class UnionOfRanges {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(UnionOfRanges.class);

        //Instantiate a workbook object.
        //Open an existing excel file.
        Workbook workbook = new Workbook(dataDir + "book1.xls");

        //Get the named ranges.
        Range[] ranges = workbook.getWorksheets().getNamedRanges();

        //Create a style object.
        Style style = workbook.createStyle();

        //Set the shading color with solid pattern type.
        style.setForegroundColor(Color.getYellow());
        style.setPattern(BackgroundType.SOLID);

        //Create a styleflag object.
        StyleFlag flag = new StyleFlag();

        //Apply the cellshading.
        flag.setCellShading(true);

        //Creates an arraylist.
        ArrayList al = new ArrayList();

        //Get the arraylist collection apply the union operation.
        al = ranges[0].union(ranges[1]);

        //Define a range object.
        Range rng;
        int frow, fcol, erow, ecol;

        for (int i = 0; i < al.size(); i++) {
            //Get a range.
            rng = (Range) al.get(i);
            frow = rng.getFirstRow();
            fcol = rng.getFirstColumn();
            erow = rng.getRowCount();
            ecol = rng.getColumnCount();

            //Apply the style to the range.
            rng.applyStyle(style, flag);

        }

        //Save the excel file.
        workbook.save(dataDir + "rngUnion.out.xls");

        // Print message
        System.out.println("Process completed successfully");
    }
}
