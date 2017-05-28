package AsposeCellsExamples.Data;

import com.aspose.cells.*;

public class CheckCustomNumberFormat {

	public static void main(String[] args) throws Exception {

		// Create a workbook
		Workbook wb = new Workbook();

		// Setting this property to true will make Aspose.Cells to throw exception
		// when invalid custom number format is assigned to Style.Custom property
		wb.getSettings().setCheckCustomNumberFormat(true);

		// Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		// Access cell A1 and put some number inside it
		Cell c = ws.getCells().get("A1");
		c.putValue(2347);

		// Access cell's style and set its Style.Custom property
		Style s = c.getStyle();

		try {
			// This line will throw exception if
			// Workbook.Settings.CheckCustomNumberFormat is set to true
			s.setCustom("ggg @ fff"); // Invalid custom number format
			c.setStyle(s);

		} 
		catch (Exception ex) {
			System.out.println("Exception Occured");
		}

		System.out.println("Done");
	}
}
