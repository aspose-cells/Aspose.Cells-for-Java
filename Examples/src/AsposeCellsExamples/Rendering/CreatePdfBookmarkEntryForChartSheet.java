package AsposeCellsExamples.Rendering;

import java.util.ArrayList;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class CreatePdfBookmarkEntryForChartSheet {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Load sample Excel file
		Workbook wb = new Workbook(srcDir + "sampleCreatePdfBookmarkEntryForChartSheet.xlsx");

		//Access all four worksheets
		Worksheet sheet1 = wb.getWorksheets().get(0);
		Worksheet sheet2 = wb.getWorksheets().get(1);
		Worksheet sheet3 = wb.getWorksheets().get(2);
		Worksheet sheet4 = wb.getWorksheets().get(3);

		//Create Pdf Bookmark Entry for Sheet1
		PdfBookmarkEntry ent1 = new PdfBookmarkEntry();
		ent1.setDestination(sheet1.getCells().get("A1"));
		ent1.setText("Bookmark-I");

		//Create Pdf Bookmark Entry for Sheet2 - Chart 
		PdfBookmarkEntry ent2 = new PdfBookmarkEntry();
		ent2.setDestination(sheet2.getCells().get("A1"));
		ent2.setText("Bookmark-II-Chart1");

		//Create Pdf Bookmark Entry for Sheet3 
		PdfBookmarkEntry ent3 = new PdfBookmarkEntry();
		ent3.setDestination(sheet3.getCells().get("A1"));
		ent3.setText("Bookmark-III");

		//Create Pdf Bookmark Entry for Sheet4 - Chart 
		PdfBookmarkEntry ent4 = new PdfBookmarkEntry();
		ent4.setDestination(sheet4.getCells().get("A1"));
		ent4.setText("Bookmark-IV-Chart2");

		//Arrange all Bookmark Entries
		ArrayList lst = new ArrayList();
		ent1.setSubEntry(lst);
		lst.add(ent2);
		lst.add(ent3);
		lst.add(ent4);

		//Create Pdf Save Options with Bookmark Entries
		PdfSaveOptions opts = new PdfSaveOptions();
		opts.setBookmark(ent1);

		//Save the output Pdf
		wb.save(outDir + "outputCreatePdfBookmarkEntryForChartSheet.pdf", opts);
		
		// Print the message
		System.out.println("CreatePdfBookmarkEntryForChartSheet executed successfully.");
	}
}
