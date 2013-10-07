/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithdata.addonfeatures.namedranges.intersectionofranges.java;

import com.aspose.cells.*;

public class IntersectionofRanges
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithdata/addonfeatures/namedranges/intersectionofranges/data/";

        //Instantiate a workbook object.
        //Open an existing excel file.
        Workbook workbook = new Workbook(dataDir + "book1.xls");

        //Get the named ranges.
        Range[] ranges = workbook.getWorksheets().getNamedRanges();

        //Check whether the first range intersect the second range.
        boolean isintersect = ranges[0].isIntersect(ranges[1]);

        //Create a style object.
        Style style = workbook.getStyles().get(workbook.getStyles().add());

        //Set the shading color with solid pattern type.
        style.setForegroundColor(Color.getYellow());
        style.setPattern(BackgroundType.SOLID);

        //Create a styleflag object.
        StyleFlag flag = new StyleFlag();

        //Apply the cellshading.
        flag.setCellShading(true);

        //If first range intersects second range.
        if (isintersect)
        {
            //Create a range by getting the intersection.
            Range intersection = ranges[0].intersect(ranges[1]);

            //Name the range.
            intersection.setName("Intersection");

            //Apply the style to the range.
            intersection.applyStyle(style, flag);

        }

        //Save the excel file.
        workbook.save(dataDir + "rngIntersection.xls");

        // Print message
        System.out.println("Process completed successfully");
    }
}




