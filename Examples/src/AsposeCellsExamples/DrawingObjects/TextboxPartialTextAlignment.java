package AsposeCellsExamples.DrawingObjects;

import com.aspose.cells.Color;
import com.aspose.cells.FillFormat;
import com.aspose.cells.GradientStyleType;
import com.aspose.cells.LineFormat;
import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.MsoLineDashStyle;
import com.aspose.cells.MsoLineStyle;
import com.aspose.cells.PdfCustomPropertiesExport;
import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.PlacementType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Shape;
import com.aspose.cells.TextBox;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class TextboxPartialTextAlignment {
	
	// The path to the documents directory.
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	
	public static void main(String[] args) throws Exception {
		// ExStart:ExportCustomPropertiesToPDF
		// Intialize an object of the Workbook class to load template file
		Workbook sourceWb = new Workbook(srcDir + "SampleTextboxExcel2016.xlsx");

		// Access the target textbox whose text is to be aligned
		Shape sourceTextBox = sourceWb.getWorksheets().get(0).getShapes().get(0);

		// Create and object of the target workbook
		Workbook destWb = new Workbook();

		// Access first worksheet from the collection
		Worksheet _sheet = destWb.getWorksheets().get(0);

		// Create new textbox
		TextBox _textBox = (TextBox)_sheet.getShapes().addShape(6,1, 0, 1, 0, 200, 200);

		// Use Html string from a template file textbox
		_textBox.setHtmlText(sourceTextBox.getHtmlText());

		// Save the workbook on disc
		destWb.save(outDir + "Output.xlsx");
		
		// Print message
		System.out.println("Textbox partial text alignment performed successfully.");
		// ExEnd:ExportCustomPropertiesToPDF		
	}
}
