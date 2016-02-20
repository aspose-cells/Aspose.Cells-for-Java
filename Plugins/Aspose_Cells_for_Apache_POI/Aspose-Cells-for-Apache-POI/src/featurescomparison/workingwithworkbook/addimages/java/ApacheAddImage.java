package featurescomparison.workingwithworkbook.addimages.java;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ApacheAddImage
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithworkbook/addimages/data/";
		
		//create a new workbook
	    Workbook wb = new XSSFWorkbook(); //or new HSSFWorkbook();

	    //add picture data to this workbook.
	    InputStream is = new FileInputStream(dataPath + "aspose.jpg");
	    byte[] bytes = IOUtils.toByteArray(is);
	    int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
	    is.close();

	    CreationHelper helper = wb.getCreationHelper();

	    //create sheet
	    Sheet sheet = wb.createSheet();

	    // Create the drawing patriarch.  This is the top level container for all shapes. 
	    Drawing drawing = sheet.createDrawingPatriarch();

	    //add a picture shape
	    ClientAnchor anchor = helper.createClientAnchor();
	    //set top-left corner of the picture,
	    //subsequent call of Picture#resize() will operate relative to it
	    anchor.setCol1(3);
	    anchor.setRow1(2);
	    Picture pict = drawing.createPicture(anchor, pictureIdx);

	    //auto-size picture relative to its top-left corner
	    pict.resize();

	    //save workbook
	    String file = dataPath + "ApacheImage.xls";
	    if(wb instanceof XSSFWorkbook) file += "x";
	    FileOutputStream fileOut = new FileOutputStream(file);
	    wb.write(fileOut);
	    fileOut.close();
	    System.out.println("Done...");
	}
}
