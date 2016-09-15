package com.aspose.cells.examples.articles;

import com.aspose.cells.ConnectionParameter;
import com.aspose.cells.ConnectionParameterCollection;
import com.aspose.cells.DBConnection;
import com.aspose.cells.ExternalConnection;
import com.aspose.cells.ExternalConnectionCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class RetrieveSQLConnectionData {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(RetrieveSQLConnectionData.class) + "articles/";
		// Create a workbook object from source file
		Workbook workbook = new Workbook(dataDir + "connection.xlsx");

		// Access the external collections
		ExternalConnectionCollection connections = workbook.getDataConnections();

		int connectionCount = connections.getCount();

		ExternalConnection connection = null;

		for (int i = 0; i < connectionCount; i++) {
			connection = connections.get(i);

			// Check if the Connection is DBConnection, then retrieve its various properties
			if (connection instanceof DBConnection) {
				DBConnection dbConn = (DBConnection) connection;

				// Retrieve DB Connection Command
				System.out.println("Command: " + dbConn.getCommand());

				// Retrieve DB Connection Command Type
				System.out.println("Command Type: " + dbConn.getCommandType());

				// Retrieve DB Connection Description
				System.out.println("Description: " + dbConn.getConnectionDescription());

				// Retrieve DB Connection ID
				System.out.println("Id: " + dbConn.getConnectionId());

				// Retrieve DB Connection Info
				System.out.println("Info: " + dbConn.getConnectionInfo());

				// Retrieve DB Connection Credentials
				System.out.println("Credentials: " + dbConn.getCredentials());

				// Retrieve DB Connection Name
				System.out.println("Name: " + dbConn.getName());

				// Retrieve DB Connection ODC File
				System.out.println("OdcFile: " + dbConn.getOdcFile());

				// Retrieve DB Connection Source File
				System.out.println("Source file: " + dbConn.getSourceFile());

				// Retrieve DB Connection Type
				System.out.println("Type: " + dbConn.getType());

				// Retrieve DB Connection Parameters Collection
				ConnectionParameterCollection parameterCollection = dbConn.getParameters();

				int paramCount = parameterCollection.getCount();

				// Iterate the Parameter Collection
				for (int j = 0; j < paramCount; j++) {

					ConnectionParameter param = parameterCollection.get(j);

					// Retrieve Parameter Cell Reference
					System.out.println("Cell reference: " + param.getCellReference());

					// Retrieve Parameter Name
					System.out.println("Parameter name: " + param.getName());

					// Retrieve Parameter Prompt
					System.out.println("Prompt: " + param.getPrompt());

					// Retrieve Parameter SQL Type
					System.out.println("SQL Type: " + param.getSqlType());

					// Retrieve Parameter Type
					System.out.println("Param Type: " + param.getType());

					// Retrieve Parameter Value
					System.out.println("Param Value: " + param.getValue());

				} // End for
			} // End if
		} // End for


	}
}
