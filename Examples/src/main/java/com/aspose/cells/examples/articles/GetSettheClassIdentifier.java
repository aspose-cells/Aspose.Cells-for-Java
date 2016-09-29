package com.aspose.cells.examples.articles;

import com.aspose.cells.OleObject;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class GetSettheClassIdentifier {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(GetSettheClassIdentifier.class) + "articles/";
		
		//Load your sample workbook which contains embedded PowerPoint ole object
		Workbook wb = new Workbook(dataDir + "sample.xls");

		//Access its first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		//Access first ole object inside the worksheet
		OleObject oleObj = ws.getOleObjects().get(0);

		//Get the class identifier of ole object in bytes and convert them into GUID
		byte[] classId = oleObj.getClassIdentifier();

		//Position of the bytes and formatting
		int[] pos = {3, 2, 1, 0, -1, 5, 4, -1, 7, 6, -1, 8, 9, -1, 10, 11, 12, 13, 14,15};

		StringBuilder sb = new StringBuilder();
		for(int i=0; i<pos.length; i++)
		{
			if(pos[i]==-1)
			{
				sb.append("-");
			}
			else
			{
				sb.append(String.format("%02X", classId[pos[i]]&0xff));
			}
		}

		//Get the GUID from conversion
		String guid = sb.toString();

		//Print the GUID
		System.out.println(guid);
	}
}
