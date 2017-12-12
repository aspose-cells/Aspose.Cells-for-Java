package AsposeCellsExamples.XmlMaps;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class FindRootElementNameOfXmlMap {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Load sample Excel file having Xml Map
		Workbook wb = new Workbook(srcDir + "sampleRootElementNameOfXmlMap.xlsx");
		  
		//Access first Xml Map inside the Workbook
		XmlMap xmap = wb.getWorksheets().getXmlMaps().get(0);
		  
		//Print Root Element Name of Xml Map on Console
		System.out.println("Root Element Name Of Xml Map: " + xmap.getRootElementName());

		// Print the message
		System.out.println("FindRootElementNameOfXmlMap executed successfully.");
	}
}
