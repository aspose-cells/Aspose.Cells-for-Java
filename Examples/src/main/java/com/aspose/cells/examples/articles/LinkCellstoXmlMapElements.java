package com.aspose.cells.examples.articles;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.XmlMap;
import com.aspose.cells.examples.Utils;

public class LinkCellstoXmlMapElements {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(LinkCellstoXmlMapElements.class) + "articles/";
		// Load sample workbook
		Workbook wb = new Workbook(dataDir + "LinkCellstoXmlMapElements_in.xlsx");

		// Access the Xml Map inside it
		XmlMap map = wb.getWorksheets().getXmlMaps().get(0);

		// Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		// Map FIELD1 and FIELD2 to cell A1 and B2
		ws.getCells().linkToXmlMap(map.getName(), 0, 0, "/root/row/FIELD1");
		ws.getCells().linkToXmlMap(map.getName(), 1, 1, "/root/row/FIELD2");

		// Map FIELD4 and FIELD5 to cell C3 and D4
		ws.getCells().linkToXmlMap(map.getName(), 2, 2, "/root/row/FIELD4");
		ws.getCells().linkToXmlMap(map.getName(), 3, 3, "/root/row/FIELD5");

		// Map FIELD7 and FIELD8 to cell E5 and F6
		ws.getCells().linkToXmlMap(map.getName(), 4, 4, "/root/row/FIELD7");
		ws.getCells().linkToXmlMap(map.getName(), 5, 5, "/root/row/FIELD8");

		// Save the workbook in xlsx format
		wb.save(dataDir + "LinkCellstoXmlMapElements_out.xlsx");
	}
}
