package featurescomparison.workingwithworksheets.reordersheetswithinworkbook.java;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

public class ApacheMoveSheetWithinWorkbook
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithworksheets/reordersheetswithinworkbook/data/";
		
		Workbook wb = new HSSFWorkbook();
	    wb.createSheet("new sheet");
	    wb.createSheet("second sheet");
	    wb.createSheet("third sheet");

	    wb.setSheetOrder("second sheet", 0);
	    wb.setSheetOrder("new sheet", 1);
	    wb.setSheetOrder("third sheet", 2);
	    
	    FileOutputStream fileOut = new FileOutputStream(dataPath + "Apache_Reordered_Out.xls");
	    wb.write(fileOut);
	    fileOut.close();
	    
	    //Print Message
        System.out.println("Reordered successfull.");
	}
}
