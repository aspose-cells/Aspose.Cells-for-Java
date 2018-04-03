package AsposeCellsExamples.DrawingObjects;

import java.io.*;
import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class AccessAndModifyLabelOfOleObject {
	
	static String srcDir = Utils.Get_SourceDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Load the sample Excel file 
		Workbook wb = new Workbook(srcDir + "sampleAccessAndModifyLabelOfOleObject.xlsx");
		 
		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);
		 
		//Access first Ole Object
		OleObject oleObject = ws.getOleObjects().get(0);
		 
		//Display the Label of the Ole Object
		System.out.println("Ole Object Label - Before: " + oleObject.getLabel());
		 
		//Modify the Label of the Ole Object
		oleObject.setLabel("Aspose APIs");

		//Save workbook to byte array output stream
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		wb.save(baos, SaveFormat.XLSX);

		//Convert output to input stream
		ByteArrayInputStream bais = new ByteArrayInputStream(baos.toByteArray());

		//Set the workbook reference to null
		wb = null;

		//Load workbook from byte array input stream
		wb = new Workbook(bais);
		 
		//Access first worksheet
		ws = wb.getWorksheets().get(0);
		 
		//Access first Ole Object
		oleObject = ws.getOleObjects().get(0);
		 
		//Display the Label of the Ole Object that has been modified earlier
		System.out.println("Ole Object Label - After: " + oleObject.getLabel());

		// Print the message
		System.out.println("AccessAndModifyLabelOfOleObject executed successfully.");
	}
}
