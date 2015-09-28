package asposefeatures.workingwithworksheets.displayandhidescrollbars.java;

import com.aspose.cells.Workbook;

public class AsposeDisplayAndHideScrollbars
{
    public static void main(String[] args) throws Exception
    {
	String dataPath = "src/asposefeatures/workingwithworksheets/displayandhidescrollbars/data/";
	
	//Instantiating a Excel object by excel file path
	Workbook workbook = new Workbook(dataPath + "book1.xls");

	//Hiding the vertical scroll bar of the Excel file
	workbook.getSettings().setVScrollBarVisible(false);

	//Hiding the horizontal scroll bar of the Excel file
	workbook.getSettings().setHScrollBarVisible(false);

	//Saving the modified Excel file in default (that is Excel 2003) format
	workbook.save(dataPath + "AsposeSrollbarsHide.xls");

	// ===============================================================
	
	//Displaying the vertical scroll bar of the Excel file
	workbook.getSettings().setVScrollBarVisible(true);

	//Displaying the horizontal scroll bar of the Excel file
	workbook.getSettings().setHScrollBarVisible(true);

	//Saving the modified Excel file in default (that is Excel 2003) format
	workbook.save(dataPath + "AsposeDisplaySrollbars.xls");
	
	System.out.println("Scrollbars. Done");
    }
}