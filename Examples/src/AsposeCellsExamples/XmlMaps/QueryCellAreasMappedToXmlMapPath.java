package AsposeCellsExamples.XmlMaps;

import java.util.*;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class QueryCellAreasMappedToXmlMapPath {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
			
		//Load sample Excel file having Xml Map
		Workbook wb = new Workbook(srcDir + "sampleXmlMapQuery.xlsx");
		  
		//Access first XML Map
		XmlMap xmap = wb.getWorksheets().getXmlMaps().get(0);
		  
		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);
		  
		//Query Xml Map from Path - /MiscData
		System.out.println("Query Xml Map from Path - /MiscData");
		ArrayList ret = ws.xmlMapQuery("/MiscData", xmap);
		  
		//Print returned ArrayList values
		for (int i = 0; i < ret.size(); i++)
		{
		    System.out.println(ret.get(i));
		}
		  
		System.out.println("");
		  
		//Query Xml Map from Path - /MiscData/row/Color
		System.out.println("Query Xml Map from Path - /MiscData/row/Color");
		ret = ws.xmlMapQuery("/MiscData/row/Color", xmap);
		  
		//Print returned ArrayList values
		for (int i = 0; i < ret.size(); i++)
		{
		    System.out.println(ret.get(i));
		}

		// Print the message
		System.out.println("QueryCellAreasMappedToXmlMapPath executed successfully.");
	}
}
