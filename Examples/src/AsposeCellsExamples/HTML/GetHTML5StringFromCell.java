package AsposeCellsExamples.HTML;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class GetHTML5StringFromCell {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Create workbook.
		Workbook wb = new Workbook();

		//Access first worksheet.
		Worksheet ws = wb.getWorksheets().get(0);

		//Access cell A1 and put some text inside it.
		Cell cell = ws.getCells().get("A1");
		cell.putValue("This is some text.");

		//Get the Normal and Html5 strings.
		String strNormal = cell.getHtmlString(false);
		String strHtml5 = cell.getHtmlString(true);

		//Print the Normal and Html5 strings on console.
		System.out.println("Normal:\r\n" + strNormal);
		System.out.println();
		System.out.println("Html5:\r\n" + strHtml5);
	 
		// Print the message
		System.out.println("GetHTML5StringFromCell executed successfully.");
	}
}
