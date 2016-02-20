package featurescomparison.workingwithcellsrowscolumns.autofitrowandcolumn.java;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ApacheAutoFit
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithcellsrowscolumns/autofitrowandcolumn/data/";
		
		InputStream inStream = new FileInputStream(dataPath + "workbook.xls");
		Workbook workbook = WorkbookFactory.create(inStream);
		
		Sheet sheet = workbook.createSheet("new sheet");
		sheet.autoSizeColumn(0); //adjust width of the first column
		sheet.autoSizeColumn(1); //adjust width of the second column
		
		FileOutputStream fileOut;
		fileOut = new FileOutputStream(dataPath + "AutoFit_Apache_Out.xls");
		workbook.write(fileOut);
		fileOut.close();
		
		System.out.println("Process Completed.");
	}
}