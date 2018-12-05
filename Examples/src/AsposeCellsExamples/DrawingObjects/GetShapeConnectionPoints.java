package AsposeCellsExamples.DrawingObjects;

import com.aspose.cells.Color;
import com.aspose.cells.FillFormat;
import com.aspose.cells.GradientStyleType;
import com.aspose.cells.LineFormat;
import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.MsoLineDashStyle;
import com.aspose.cells.MsoLineStyle;
import com.aspose.cells.Oval;
import com.aspose.cells.PdfCustomPropertiesExport;
import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.PlacementType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Shape;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.b.a.b.zo;

import AsposeCellsExamples.Utils;

public class GetShapeConnectionPoints {

	public static void main(String[] args) throws Exception {

        // ExStart:1
        // Instantiate a new Workbook.
        Workbook workbook = new Workbook();

        // Get the first worksheet in the book.
        Worksheet worksheet = workbook.getWorksheets().get(0);

        // Add a new textbox to the collection.
        int textboxIndex = worksheet.getTextBoxes().add(2, 1, 160, 200);

        // Access your text box which is also a shape object from shapes collection
        Shape shape = workbook.getWorksheets().get(0).getShapes().get(0);

        // Get all the connection points in this shape
        float[][] ConnectionPoints = shape.getConnectionPoints();

        // Display all the shape points
        for (float[] pt : ConnectionPoints)
        {
            System.out.println(pt[0]);
            System.out.println(pt[1]);
        }
        // ExEnd:1
		
		// Print message
		System.out.println("GetShapeConnectionPoints executed successfully.");
	}

}
