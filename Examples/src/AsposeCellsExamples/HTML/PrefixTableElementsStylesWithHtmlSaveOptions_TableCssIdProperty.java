package AsposeCellsExamples.HTML;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class PrefixTableElementsStylesWithHtmlSaveOptions_TableCssIdProperty {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());

		//Create workbook object
		Workbook wb = new Workbook();

		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		//Access cell B5 and put value inside it
		Cell cell = ws.getCells().get("B5");
		cell.putValue("This is some text.");

		//Set the style of the cell - font color is Red
		Style st = cell.getStyle();
		st.getFont().setColor(Color.getRed());
		cell.setStyle(st);

		//Specify html save options - specify table css id
		HtmlSaveOptions opts = new HtmlSaveOptions();
		opts.setTableCssId("MyTest_TableCssId");

		//Save the workbook in html 
		wb.save(outDir + "outputTableCssId.html", opts);

		// Print the message
		System.out.println("PrefixTableElementsStylesWithHtmlSaveOptions_TableCssIdProperty executed successfully.");
	}
}
