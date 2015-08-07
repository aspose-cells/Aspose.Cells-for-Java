package featurescomparison.workingwithcellsrowscolumns.addcomments.java;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ApacheAddCommentsToCell 
{
	public static void main(String[] args) throws IOException 
	{
		String dataPath = "src/featurescomparison/workingwithcellsrowscolumns/addcomments/data/";
		
		Workbook wb = new XSSFWorkbook(); //or new HSSFWorkbook();

	    CreationHelper factory = wb.getCreationHelper();

	    Sheet sheet = wb.createSheet();
	    
	    Row row   = sheet.createRow(3);
	    Cell cell = row.createCell(5);
	    cell.setCellValue("F4");
	    
	    Drawing drawing = sheet.createDrawingPatriarch();

	    // When the comment box is visible, have it show in a 1x3 space
	    ClientAnchor anchor = factory.createClientAnchor();
	    anchor.setCol1(cell.getColumnIndex());
	    anchor.setCol2(cell.getColumnIndex()+1);
	    anchor.setRow1(row.getRowNum());
	    anchor.setRow2(row.getRowNum()+3);

	    // Create the comment and set the text+author
	    Comment comment = drawing.createCellComment(anchor);
	    RichTextString str = factory.createRichTextString("Hello, World!");
	    comment.setString(str);
	    comment.setAuthor("Apache POI");

	    // Assign the comment to the cell
	    cell.setCellComment(comment);

	    String fname = "AsposeComment-xssf.xls";
	    if(wb instanceof XSSFWorkbook) fname += "x";
	    FileOutputStream out = new FileOutputStream(dataPath + fname);
	    wb.write(out);
	    out.close();
	    System.out.println("Done.");
	}
}
