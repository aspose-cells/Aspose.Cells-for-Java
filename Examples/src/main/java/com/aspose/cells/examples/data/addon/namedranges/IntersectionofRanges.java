package com.aspose.cells.examples.data.addon.namedranges;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class IntersectionofRanges {

	public static void main(String[] args) throws Exception {
		// ExStart:IntersectionofRanges
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(IntersectionofRanges.class);

		// Instantiate a workbook object.
		// Open an existing excel file.
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Get the named ranges.
		Range[] ranges = workbook.getWorksheets().getNamedRanges();

		// Check whether the first range intersect the second range.
		boolean isintersect = ranges[0].isIntersect(ranges[1]);

		// Create a style object.
		Style style = workbook.createStyle();

		// Set the shading color with solid pattern type.
		style.setForegroundColor(Color.getYellow());
		style.setPattern(BackgroundType.SOLID);

		// Create a styleflag object.
		StyleFlag flag = new StyleFlag();

		// Apply the cellshading.
		flag.setCellShading(true);

		// If first range intersects second range.
		if (isintersect) {
			// Create a range by getting the intersection.
			Range intersection = ranges[0].intersect(ranges[1]);

			// Name the range.
			intersection.setName("Intersection");

			// Apply the style to the range.
			intersection.applyStyle(style, flag);

		}

		// Save the excel file.
		workbook.save(dataDir + "rngIntersection.out.xls");

		// Print message
		System.out.println("Process completed successfully");
		// ExEnd:IntersectionofRanges
	}
}
