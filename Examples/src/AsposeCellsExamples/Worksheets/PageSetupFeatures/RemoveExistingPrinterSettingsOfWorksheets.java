package AsposeCellsExamples.Worksheets.PageSetupFeatures;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class RemoveExistingPrinterSettingsOfWorksheets {
	public static void main(String[] args) throws Exception {
		String srcDir = Utils.Get_SourceDirectory();
		String outDir = Utils.Get_OutputDirectory();

		//Load source Excel file
		Workbook wb = new Workbook(srcDir + "sampleRemoveExistingPrinterSettingsOfWorksheets.xlsx");
		 
		//Get the sheet counts of the workbook
		int sheetCount = wb.getWorksheets().getCount();
		 
		//Iterate all sheets
		for(int i=0; i<sheetCount; i++)
		{
		    //Access the i-th worksheet
		    Worksheet ws = wb.getWorksheets().get(i);
		 
		    //Access worksheet page setup
		    PageSetup ps = ws.getPageSetup();
		 
		    //Check if printer settings for this worksheet exist
		    if(ps.getPrinterSettings() != null)
		    {
		        //Print the following message
		        System.out.println("PrinterSettings of this worksheet exist.");
		 
		        //Print sheet name and its paper size
		        System.out.println("Sheet Name: " + ws.getName());
		        System.out.println("Paper Size: " + ps.getPaperSize());
		 
		        //Remove the printer settings by setting them null
		        ps.setPrinterSettings(null);
		        System.out.println("Printer settings of this worksheet are now removed by setting it null.");
		        System.out.println("");
		    }//if
		}//for
		 
		//Save the workbook
		wb.save(outDir + "outputRemoveExistingPrinterSettingsOfWorksheets.xlsx");

	}
}
