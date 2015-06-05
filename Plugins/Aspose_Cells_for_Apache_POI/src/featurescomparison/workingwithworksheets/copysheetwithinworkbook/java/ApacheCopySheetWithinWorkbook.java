package featurescomparison.workingwithworksheets.copysheetwithinworkbook.java;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ApacheCopySheetWithinWorkbook
{
	public static void main(String[] args) throws Exception
	{
		Workbook wb = new HSSFWorkbook();
	    wb.createSheet("new sheet");
	    wb.createSheet("second sheet");
	    Sheet cloneSheet = wb.cloneSheet(0);
	    
	    // now you have to manually copy all the data into new sheet from the cloneSheet
	    
	    //Print Message
        System.out.println("Cloned successfull.");
	}
}