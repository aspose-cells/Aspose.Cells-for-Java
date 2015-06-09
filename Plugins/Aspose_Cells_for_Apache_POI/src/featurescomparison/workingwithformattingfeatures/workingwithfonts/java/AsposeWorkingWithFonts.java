package featurescomparison.workingwithformattingfeatures.workingwithfonts.java;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.Font;
import com.aspose.cells.FontUnderlineType;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class AsposeWorkingWithFonts
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithformattingfeatures/workingwithfonts/data/";
		
		//Instantiating a Workbook object
		Workbook workbook = new Workbook();

		//Accessing the worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);
		Cells cells = worksheet.getCells();

		//Adding some value to cell
		Cell cell = cells.get("A1");
		cell.setValue("This is Aspose test of fonts!");

		//Setting the font name to "Times New Roman"
		Style style = cell.getStyle();
		Font font = style.getFont();
		font.setName("Courier New");
		font.setSize(24);
		font.setBold(true);
		font.setUnderline(FontUnderlineType.SINGLE);
		font.setColor(Color.getBlue());
		font.setStrikeout(true);
		//font.setSubscript(true);

		cell.setStyle(style); 

		//Saving the modified Excel file in default format
		workbook.save(dataPath + "AsposeFonts.xls");
		
		System.out.println("Aspose Fonts Created.");
	}
}
