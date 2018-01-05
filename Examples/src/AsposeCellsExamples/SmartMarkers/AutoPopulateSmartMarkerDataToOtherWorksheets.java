package AsposeCellsExamples.SmartMarkers;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class AutoPopulateSmartMarkerDataToOtherWorksheets {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		// Create Connection object - connect to Microsoft Access Students Database
		java.sql.Connection conn = java.sql.DriverManager.getConnection("jdbc:ucanaccess://" + srcDir + "sampleAutoPopulateSmartMarkerDataToOtherWorksheets.accdb");

		// Create SQL Statement with Connection object
		java.sql.Statement st = conn.createStatement();

		// Execute SQL Query and obtain ResultSet
		java.sql.ResultSet rsEmployees = st.executeQuery("SELECT * FROM Employees");
		
		//Create empty workbook
		Workbook wb = new Workbook();

		//Access first worksheet and add smart marker in cell A1
		Worksheet ws = wb.getWorksheets().get(0);
		ws.getCells().get("A1").putValue("&=Employees.EmployeeID");

		//Add second worksheet and add smart marker in cell A1
		wb.getWorksheets().add();
		ws = wb.getWorksheets().get(1);
		ws.getCells().get("A1").putValue("&=Employees.EmployeeID");

		//Create workbook designer
		WorkbookDesigner wd = new WorkbookDesigner(wb);

		//Set data source with result set
		wd.setDataSource("Employees", rsEmployees, 15);

		//Process smart marker tags in first and second worksheet
		wd.process(0, false);
		wd.process(1, false);

		//Save the workbook
		wb.save(outDir + "outputAutoPopulateSmartMarkerDataToOtherWorksheets.xlsx");

		// Print the message
		System.out.println("AutoPopulateSmartMarkerDataToOtherWorksheets executed successfully.");
	}
}