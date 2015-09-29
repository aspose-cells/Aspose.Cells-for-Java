package featurescomparison.workingwithdata.hyperlink.java;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.FontUnderlineType;
import com.aspose.cells.HyperlinkCollection;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

public class AsposeHyperlinks
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithdata/hyperlink/data/";
		
		//Instantiating a Workbook object
		Workbook workbook = new Workbook();

		//Obtaining the reference of the first worksheet.
		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet sheet = worksheets.get(0);
		HyperlinkCollection hyperlinks = sheet.getHyperlinks();

		//Adding a hyperlink to a URL at "A1" cell
		hyperlinks.add("A1",1,1,"http://www.aspose.com");

		//============ Link to Cell =================
		//Setting a value to the "A1" cell
		Cells cells = sheet.getCells();
		Cell cell = cells.get("A2");
		cell.setValue("Link to B9");

		setFormatting(cell);

		hyperlinks = sheet.getHyperlinks();

		//Adding an internal hyperlink to the "B9" cell of the other worksheet "Sheet1" in
		//the same Excel file

		hyperlinks.add("A2",1 ,1, "Sheet1!B9");
		
		//============ Link to External File ========
		
		cell = cells.get("A3");
		cell.setValue("External Link");

		setFormatting(cell);
		
		hyperlinks = sheet.getHyperlinks();

		//Adding a link to the external file
		hyperlinks.add("A3", 1, 1, "book1.xls");

		//Saving the Excel file
		//workbook.save("c:\\book2.xls");
		workbook.save(dataPath + "AsposeHyperlink.xls");
		
		System.out.println("Done ...");
	}
	//=============================================================
	private static void setFormatting(Cell cell)
	{
		//Setting the font color of the cell to Blue
		Style style = cell.getStyle();
		style.getFont().setColor(Color.getBlue());

		//Setting the font of the cell to Single Underline
		style.getFont().setUnderline(FontUnderlineType.SINGLE);
		cell.setStyle(style);		


	
	
	}
}
