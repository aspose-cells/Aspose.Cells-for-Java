package com.aspose.cells.examples.articles;

import com.aspose.cells.examples.Utils;
import com.aspose.cells.*;

public class ExportXmlDataOfXmlMap {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ExportXmlDataOfXmlMap.class) + "articles/";
		
		//Load source workbook
		Workbook wb = new Workbook(dataDir + "sample_Export-Xml-Data-linked.xlsx");

		//Export all XML data from all XML Maps inside the Workbook
		for (int i = 0; i < wb.getWorksheets().getXmlMaps().getCount(); i++)
		{
		    //Access the XML Map
		    XmlMap map = wb.getWorksheets().getXmlMaps().get(i);

		    //Exports its XML Data
		    wb.exportXml(map.getName(), dataDir +  map.getName() + ".xml");
		}
	}
}