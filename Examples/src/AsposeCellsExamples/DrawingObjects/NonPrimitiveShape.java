package AsposeCellsExamples.DrawingObjects;

import com.aspose.cells.AutoShapeType;
import com.aspose.cells.Shape;
import com.aspose.cells.ShapePath;
import com.aspose.cells.ShapePathCollection;
import com.aspose.cells.ShapePathPoint;
import com.aspose.cells.ShapePathPointCollection;
import com.aspose.cells.ShapeSegmentPath;
import com.aspose.cells.ShapeSegmentPathCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class NonPrimitiveShape {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the resource directory
		String dataDir = Utils.getSharedDataDir(NonPrimitiveShape.class) + "DrawingObjects/";

		Workbook workbook = new Workbook(dataDir + "NonPrimitiveShape.xlsx");

		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Accessing the user defined shape
		Shape shape = worksheet.getShapes().get(0);

		if (shape.getAutoShapeType() == AutoShapeType.NOT_PRIMITIVE) {

			// Access Shape paths
			ShapePathCollection shapePathCollection = shape.getPaths();

			// Access information of individual shape path
			ShapePath shapePath = shapePathCollection.get(0);

			// Access shape segment path list
			ShapeSegmentPathCollection shapeSegmentPathCollection = shapePath.getPathSegementList();

			// Access individual segment path
			ShapeSegmentPath shapeSegmentPath = shapeSegmentPathCollection.get(0);
			
			ShapePathPointCollection segmentPoints = shapeSegmentPath.getPoints();
			
			for (Object obj : segmentPoints) {
				ShapePathPoint pathPoint = (ShapePathPoint) obj;
                System.out.println("X: " + pathPoint.getX() + ", Y: " + pathPoint.getY());
            }
		}
        // ExEnd:1
	}
}