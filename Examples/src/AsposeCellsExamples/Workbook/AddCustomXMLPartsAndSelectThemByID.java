package AsposeCellsExamples.Workbook;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class AddCustomXMLPartsAndSelectThemByID { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		// Create empty workbook.
		Workbook wb = new Workbook();

		// Some data in the form of byte array.
		// Please use correct XML and Schema instead.
		byte[] btsData = new byte[] { 1, 2, 3 };
		byte[] btsSchema = new byte[] { 1, 2, 3 };

		// Create four custom xml parts.
		wb.getCustomXmlParts().add(btsData, btsSchema);
		wb.getCustomXmlParts().add(btsData, btsSchema);
		wb.getCustomXmlParts().add(btsData, btsSchema);
		wb.getCustomXmlParts().add(btsData, btsSchema);

		// Assign ids to custom xml parts.
		wb.getCustomXmlParts().get(0).setID("Fruit");
		wb.getCustomXmlParts().get(1).setID("Color");
		wb.getCustomXmlParts().get(2).setID("Sport");
		wb.getCustomXmlParts().get(3).setID("Shape");

		// Specify search custom xml part id.
		String srchID = "Fruit";
		srchID = "Color";
		srchID = "Sport";

		// Search custom xml part by the search id.
		CustomXmlPart cxp = wb.getCustomXmlParts().selectByID(srchID);

		// Print the found or not found message on console.
		if (cxp == null)
		{
			System.out.println("Not Found: CustomXmlPart ID " + srchID);
		}
		else
		{
			System.out.println("Found: CustomXmlPart ID " + srchID);
		}
		
		// Print the message
		System.out.println("AddCustomXMLPartsAndSelectThemByID executed successfully.");
	}
}
