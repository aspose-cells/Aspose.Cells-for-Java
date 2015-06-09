package featurescomparison.workingwithcellsrowscolumns.hideunhidecells.java;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ApacheHideUnHideCells
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithcellsrowscolumns/hideunhidecells/data/";
		
		InputStream inStream = new FileInputStream(dataPath + "workbook.xls");
		Workbook workbook = WorkbookFactory.create(inStream);
		Sheet sheet = workbook.createSheet();
		Row row = sheet.createRow(0);
		row.setZeroHeight(true);
		
		FileOutputStream fileOut = new FileOutputStream(dataPath + "hideUnhideCells_Apache_Out.xls");
		workbook.write(fileOut);
		fileOut.close();
		
		System.out.println("Process Completed.");
	}
}
