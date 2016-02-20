package featurescomparison.workingwithworkbook.createnewworkbook.java;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

public class ApachePOINewWorkbook
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithworkbook/createnewworkbook/data/";
		
		Workbook wb = new HSSFWorkbook();

		FileOutputStream fileOut;
		fileOut = new FileOutputStream(dataPath + "newWorkBook_Apache_Out.xls");
		wb.write(fileOut);
		fileOut.close();

		System.out.println("File Created.");
	}
}