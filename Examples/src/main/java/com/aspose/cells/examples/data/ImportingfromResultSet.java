package com.aspose.cells.examples.data;

import java.beans.Statement;
import java.sql.DriverManager;
import java.sql.ResultSet;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ImportingfromResultSet {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ImportingfromResultSet.class) + "data/";
		// Create a new Workbook.
		Workbook workbook = new Workbook();

		// Define the Access Database URL String constant.
		final String DB_URL = "jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ=D:\\test\\Northwind.mdb";

		// Load the JDBC-ODBC bridge driver.
		Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");

		// Define the connection.
		Connection conn = (Connection) DriverManager.getConnection(DB_URL);

		// Create the Statement with the specified cursor type and lock option.
		Statement stmt = conn.createStatement(ResultSet.TYPE_SCROLL_SENSITIVE, ResultSet.CONCUR_READ_ONLY);

		// Get the ResultSet executing the SQL statement.
		ResultSet rs = stmt.executeQuery("select EmployeeID,LastName,FirstName,Title,City from Employees");

		// Fetch the first worksheet.
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Import the ResultSet to the worksheet.
		worksheet.getCells().importResultSet(rs, "A1", true);

		// Save the excel file.
		workbook.save(dataDir + "ImportingfromResultSet_out.xls");
	}
}
