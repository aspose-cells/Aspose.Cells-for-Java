package featurescomparison.workingwithcellsrowscolumns.iterate.java;

import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ApacheIterateRowsnCols
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithcellsrowscolumns/iterate/data/";
		
		InputStream inStream = new FileInputStream(dataPath + "workbook.xls");
		Workbook wb = WorkbookFactory.create(inStream);
		Sheet sheet = wb.getSheetAt(0);
	    for (Row row : sheet) 
	    {
	      for (Cell cell : row) 
	      {
	        System.out.println("Iteration.");
	      }
	    }
	}
}