package AsposeCellsExamples.DrawingObjects;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class SpecifyFarEastAndLatinNameOfFontInTextOptionsOfShape {
	
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		// Create empty workbook.
		Workbook wb = new Workbook();
		 
		// Access first worksheet.
		Worksheet ws = wb.getWorksheets().get(0);
		 
		// Add textbox inside the worksheet.
		int idx = ws.getTextBoxes().add(5, 5, 50, 200);
		TextBox tb = ws.getTextBoxes().get(idx);
		 
		// Set the text of the textbox.
		tb.setText("こんにちは世界");
		 
		// Specify the Far East and Latin name of the font.
		//tb.getTextOptions().setLatinName("Comic Sans MS");
		//tb.getTextOptions().setFarEastName("KaiTi");
		 
		// Save the output Excel file.
		wb.save(outDir + "outputSpecifyFarEastAndLatinNameOfFontInTextOptionsOfShape.xlsx", SaveFormat.XLSX);
		 
		// Print the message
		System.out.println("SpecifyFarEastAndLatinNameOfFontInTextOptionsOfShape executed successfully.");
	}
}
