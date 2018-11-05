package AsposeCellsExamples.XmlMaps;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class GetXMLPathFromListObjectTable {
	
	// The path to the documents directory.
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		// ExStart:1
        // Load XLSX file containing data from XML file
        Workbook workbook = new Workbook(srcDir +   "XML Data.xlsx");

        // Access the first worksheet
        Worksheet ws = workbook.getWorksheets().get(0);

        // Access ListObject from the first sheet
        ListObject listObject = ws.getListObjects().get(0);

        // Get the url of the list object's xml map data binding
        String url = listObject.getXmlMap().getDataBinding().getUrl();

        // Display XML file name
        System.out.println(url);
		// ExEnd:1		
		
		// Print message
		System.out.println("Get XML Path From List Object/Table performed successfully.");
	}
}
