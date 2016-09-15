package com.aspose.cells.examples.articles;

import com.aspose.cells.ConnectionParameter;
import com.aspose.cells.ConnectionParameterCollection;
import com.aspose.cells.DBConnection;
import com.aspose.cells.ExternalConnection;
import com.aspose.cells.ExternalConnectionCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class RetrievingExternalDataSourcesDetails {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(RetrievingExternalDataSourcesDetails.class) + "articles/";
		// Open the template Excel file
		Workbook workbook = new Workbook(dataDir + "connection.xlsx");

		// Get the external data connections
		ExternalConnectionCollection connections = workbook.getDataConnections();
		// Get the count of the collection connection
		int connectionCount = connections.getCount();

		// Create an external connection object
		ExternalConnection connection = null;

		// Loop through all the connections in the file
		for (int i = 0; i < connectionCount; i++) {
			connection = connections.get(i);
			if (connection instanceof DBConnection) {
				// Instantiate the DB Connection
				DBConnection dbConn = (DBConnection) connection;

				// Print the complete details of the object
				System.out.println("Command: " + dbConn.getCommand());
				System.out.println("Command Type: " + dbConn.getCommandType());
				System.out.println("Description: " + dbConn.getConnectionDescription());
				System.out.println("Id: " + dbConn.getConnectionId());
				System.out.println("Info: " + dbConn.getConnectionInfo());
				System.out.println("Credentials: " + dbConn.getCredentials());
				System.out.println("Name: " + dbConn.getName());
				System.out.println("OdcFile: " + dbConn.getOdcFile());
				System.out.println("Source file: " + dbConn.getSourceFile());
				System.out.println("Type: " + dbConn.getType());

				// Get the parameters collection (if the connection object has)
				ConnectionParameterCollection parameterCollection = dbConn.getParameters();
				// Loop through all the parameters and obtain the details
				int paramCount = parameterCollection.getCount();
				for (int j = 0; j < paramCount; j++) {
					ConnectionParameter param = parameterCollection.get(j);
					System.out.println("Cell reference: " + param.getCellReference());
					System.out.println("Parameter name: " + param.getName());
					System.out.println("Prompt: " + param.getPrompt());
					System.out.println("SQL Type: " + param.getSqlType());
					System.out.println("Param Type: " + param.getType());
					System.out.println("Param Value: " + param.getValue());
				}
			}
		}

	}
}
