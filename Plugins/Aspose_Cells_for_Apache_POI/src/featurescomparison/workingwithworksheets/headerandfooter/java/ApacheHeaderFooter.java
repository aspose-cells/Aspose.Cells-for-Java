package featurescomparison.workingwithworksheets.headerandfooter.java;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ApacheHeaderFooter
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithworksheets/headerandfooter/data/";
		
		Workbook wb = new HSSFWorkbook();
	    Sheet sheet = wb.createSheet("new sheet");

	    Header header = sheet.getHeader();
	    header.setCenter("Center Header");
	    header.setLeft("Left Header");
	    header.setRight(HSSFHeader.font("Stencil-Normal", "Italic") +
	                    HSSFHeader.fontSize((short) 16) + "Right w/ Stencil-Normal Italic font and size 16");

	    FileOutputStream fileOut = new FileOutputStream(dataPath + "ApacheHeaderFooter_Out.xls");
	    wb.write(fileOut);
	    fileOut.close();
	    
	    System.out.println("Aspose Header Footer Created.");
	}
}
