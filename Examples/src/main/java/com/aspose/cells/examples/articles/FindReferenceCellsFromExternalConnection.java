package com.aspose.cells.examples.articles;

import com.aspose.cells.CellsHelper;
import com.aspose.cells.ExternalConnection;
import com.aspose.cells.ListObject;
import com.aspose.cells.Name;
import com.aspose.cells.QueryTable;
import com.aspose.cells.Range;
import com.aspose.cells.TableDataSourceType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class FindReferenceCellsFromExternalConnection {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory
		String dataDir = Utils.getSharedDataDir(FindReferenceCellsFromExternalConnection.class) + "articles/";

		// Load workbook object
		Workbook workbook = new Workbook(dataDir + "sample.xlsm");

		// Check all the connections inside the workbook
		for (int i = 0; i < workbook.getDataConnections().getCount(); i++) {
			ExternalConnection externalConnection = workbook.getDataConnections().get(i);
			System.out.println("connection: " + externalConnection.getName());
			PrintTables(workbook, externalConnection);
			System.out.println();
		}
	}

	public static void PrintTables(Workbook workbook, ExternalConnection ec) {
		// Iterate all the worksheets
		for (int j = 0; j < workbook.getWorksheets().getCount(); j++) {
			Worksheet worksheet = workbook.getWorksheets().get(j);

			// Check all the query tables in a worksheet
			for (int k = 0; k < worksheet.getQueryTables().getCount(); k++) {
				QueryTable qt = worksheet.getQueryTables().get(k);

				// Check if query table is related to this external connection
				if (ec.getId() == qt.getConnectionId() && qt.getConnectionId() >= 0) {
					// Print the query table name and print its "Refers To"
					// range
					System.out.println("querytable " + qt.getName());
					String n = qt.getName().replace('+', '_').replace('=', '_');
					Name name = workbook.getWorksheets().getNames().get("'" + worksheet.getName() + "'!" + n);
					if (name != null) {
						Range range = name.getRange();
						if (range != null) {
							System.out.println("Refers To: " + range.getRefersTo());
						}
					}
				}
			}

			// Iterate all the list objects in this worksheet
			for (int k = 0; k < worksheet.getListObjects().getCount(); k++) {
				ListObject table = worksheet.getListObjects().get(k);

				// Check the data source type if it is query table
				if (table.getDataSourceType() == TableDataSourceType.QUERY_TABLE) {
					// Access the query table related to list object
					QueryTable qt = table.getQueryTable();

					// Check if query table is related to this external
					// connection
					if (ec.getId() == qt.getConnectionId() && qt.getConnectionId() >= 0) {
						// Print the query table name and print its refersto
						// range
						System.out.println("querytable " + qt.getName());
						System.out.println("Table " + table.getDisplayName());
						System.out.println("refersto: " + worksheet.getName() + "!"
								+ CellsHelper.cellIndexToName(table.getStartRow(), table.getStartColumn()) + ":"
								+ CellsHelper.cellIndexToName(table.getEndRow(), table.getEndColumn()));
					}
				}
			}
		}
	}// end-PrintTables

}
