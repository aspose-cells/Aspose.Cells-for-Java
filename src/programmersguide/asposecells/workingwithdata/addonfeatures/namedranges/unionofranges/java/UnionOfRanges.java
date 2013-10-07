/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithdata.addonfeatures.namedranges.unionofranges.java;

import com.aspose.cells.*;

import java.util.ArrayList;

public class UnionOfRanges
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithdata/addonfeatures/namedranges/unionofranges/data/";

        //Instantiate a workbook object.
        //Open an existing excel file.
        Workbook workbook = new Workbook(dataDir + "book1.xls");

        //Get the named ranges.
        Range[] ranges = workbook.getWorksheets().getNamedRanges();

        //Create a style object.
        Style style = workbook.getStyles().get(workbook.getStyles().add());

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

        for (int i = 0; i < al.size(); i++)
        {
            //Get a range.
            rng = (Range)al.get(i);
            frow = rng.getFirstRow();
            fcol = rng.getFirstColumn();
            erow = rng.getRowCount();
            ecol = rng.getColumnCount();

            //Apply the style to the range.
            rng.applyStyle(style, flag);

        }

        //Save the excel file.
        workbook.save(dataDir + "rngUnion.xls");

        // Print message
        System.out.println("Process completed successfully");
    }
}




