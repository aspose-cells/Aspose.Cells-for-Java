package AsposeCellsExamples.Data;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ImportDataFromMicrosoftAccessDatabaseResultSetObjectToWorksheet {

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());

		String srcDir = Utils.Get_SourceDirectory();
		String outDir = Utils.Get_OutputDirectory();

		// Create Connection object - connect to Microsoft Access Students
		// Database
		java.sql.Connection conn = java.sql.DriverManager
				.getConnection("jdbc:ucanaccess://" + srcDir + "Students.accdb");

		// Create SQL Statement with Connection object
		java.sql.Statement st = conn.createStatement();

		// Execute SQL Query and obtain ResultSet
		java.sql.ResultSet rs = st.executeQuery("SELECT * FROM Student");

		// Create workbook object
		Workbook wb = new Workbook();

		// Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		// Access cells collection
		Cells cells = ws.getCells();

		// Create import table options
		ImportTableOptions options = new ImportTableOptions();

		// Import Result Set at (row=2, column=2)
		cells.importResultSet(rs, 2, 2, options);

		// Execute SQL Query and obtain ResultSet again
		rs = st.executeQuery("SELECT * FROM Student");

		// Import Result Set at cell G10
		cells.importResultSet(rs, "G10", options);

		// Autofit columns
		ws.autoFitColumns();

		// Save the workbook
		wb.save(outDir + "outputImportResultSet.xlsx");

		// Print the message
		System.out.println("ImportDataFromMicrosoftAccessDatabaseResultSetObjectToWorksheet executed successfully.");
	}
}
