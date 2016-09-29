package com.aspose.cells.examples.articles;

import com.aspose.cells.DBConnection;
import com.aspose.cells.ExternalConnection;
import com.aspose.cells.OLEDBCommandType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ModifyExistingSQLDataConnection {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ModifyExistingSQLDataConnection.class) + "articles/";
		// Create a workbook object from source file
		Workbook workbook = new Workbook(dataDir + "DataConnection.xlsx");

		// Access first Data Connection
		ExternalConnection conn = workbook.getDataConnections().get(0);

		// Change the Data Connection Name and Odc file
		conn.setName("MyConnectionName");
		conn.setOdcFile(dataDir + "MyDefaulConnection.odc");

		// Change the Command Type, Command and Connection String
		DBConnection dbConn = (DBConnection) conn;
		dbConn.setCommandType(OLEDBCommandType.SQL_STATEMENT);
		dbConn.setCommand("Select * from AdminTable");
		dbConn.setConnectionInfo(
				"Server=myServerAddress;Database=myDataBase;User ID=myUsername;Password=myPassword;Trusted_Connection=False");

		// Save the workbook
		workbook.save(dataDir + "MESQLDataConnection_out.xlsx");

	}
}
