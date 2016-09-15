package com.aspose.cells.examples.articles;

import java.io.Console;

import com.aspose.cells.ExternalConnection;
import com.aspose.cells.WebQueryConnection;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class WorkingWithExternalDataConnection {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(WorkingWithExternalDataConnection.class) + "articles/";
		
		Workbook workbook = new Workbook(dataDir + "WebQuerySample.xlsx");

		ExternalConnection connection = workbook.getDataConnections().get(0);

		if (connection instanceof WebQueryConnection)
		{
		    WebQueryConnection webQuery = (WebQueryConnection)connection;
		    System.out.println("Web Query URL: " + webQuery.getUrl());
		}

	}
}
