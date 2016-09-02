package com.aspose.cells.examples.DrawingObjects;

import java.util.ArrayList;

import com.aspose.cells.AutoShapeType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class NonPrimitiveShape {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(NonPrimitiveShape.class);

		Workbook workbook = new Workbook(dataDir + "book1.xls");

		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Accessing the user defined shape
		com.aspose.cells.Shape shape = worksheet.getShapes().get(0);

		if (shape.getAutoShapeType() == AutoShapeType.NOT_PRIMITIVE) {
			// Access shape's data
			com.aspose.cells.GeomPathsInfo geomPathsInfo = shape.getPathsInfo();

			// Access path list
			ArrayList<com.aspose.cells.GeomPathInfo> pathList = geomPathsInfo.getPathList();

			// Access information of indvidual path info
			GeomPathInfo pathInfo = pathList.get(0);

			// Access path segment list
			ArrayList<MsoPathInfo> pathSegments = pathInfo.getPathSegementList();

			// Access individual path segment
			MsoPathInfo pathSegment = pathSegments.get(0);

			// Gets the points in path segment
			ArrayList segmentPoints = pathSegment.getPointList();
		}
	}
}
