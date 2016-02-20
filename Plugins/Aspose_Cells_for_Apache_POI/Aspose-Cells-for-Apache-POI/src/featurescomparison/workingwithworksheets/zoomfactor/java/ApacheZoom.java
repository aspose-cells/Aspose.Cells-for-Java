package featurescomparison.workingwithworksheets.zoomfactor.java;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ApacheZoom
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithworksheets/zoomfactor/data/";
		
		Workbook wb = new HSSFWorkbook();
	    Sheet sheet1 = wb.createSheet("new sheet");
	    sheet1.setZoom(3,4);   // 75 percent magnification
	    
	    // Write the output to a file
	    FileOutputStream fileOut = new FileOutputStream(dataPath + "ApacheZoom_Out.xls");
	    wb.write(fileOut);
	    fileOut.close();	

	    System.out.println("Process Completed Successfully.");
	}
}
