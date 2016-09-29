package com.aspose.cells.examples.articles;

import com.aspose.cells.Scenario;
import com.aspose.cells.ScenarioInputCellCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class CreateScenariosfromWorksheets {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CreateScenariosfromWorksheets.class) + "articles/";
		// Instantiate the Workbook
		// Load an Excel file
		Workbook workbook = new Workbook(dataDir + "Bk_scenarios.xlsx");

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Remove the existing first scenario from the sheet
		worksheet.getScenarios().removeAt(0);

		// Create a scenario
		int i = worksheet.getScenarios().add("MyScenario");
		// Get the scenario
		Scenario scenario = worksheet.getScenarios().get(i);
		// Add comment to it
		scenario.setComment("Test sceanrio is created.");
		// Get the input cells for the scenario
		ScenarioInputCellCollection sic = scenario.getInputCells();
		// Add the scenario on B4 (as changing cell) with default value
		sic.add(3, 1, "1100000");

		// Save the Excel file.
		workbook.save(dataDir + "CSfromWorksheets_out.xlsx");

	}
}
