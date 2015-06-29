package asposefeatures.workingwithworksheets.displayandhidetabs.java;

import com.aspose.cells.Workbook;

public class AsposeDisplayAndHideTabs
{
    public static void main(String[] args) throws Exception
    {
	String dataPath = "src/asposefeatures/workingwithworksheets/displayandhidetabs/data/";
	
	//Instantiating a Workbook object by excel file path
	Workbook workbook = new Workbook(dataPath + "book1.xls");

	//Hiding the tabs of the Excel file
	workbook.getSettings().setShowTabs(false);

	//Saving the modified Excel file in default (that is Excel 2003) format
	workbook.save(dataPath + "AsposeHideTabs.xls");

	// ===============================================================
	
	//Displaying the tabs of the Excel file
	workbook.getSettings().setShowTabs(true);

	//Saving the modified Excel file in default (that is Excel 2003) format
	workbook.save(dataPath + "AsposeDisplayTabs.xls");
	
	System.out.println("Tabs. Done");
    }
}