package AsposeCellsExamples.Rendering;

import java.util.*;
import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class AddPDFBookmarksWithNamedDestinations { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();


	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
	
		//Load source Excel file
		Workbook wb = new Workbook(srcDir + "samplePdfBookmarkEntry_DestinationName.xlsx");
		  
		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);
		  
		//Access cell C5
		Cell cell = ws.getCells().get("C5");
		  
		//Create Bookmark and Destination for this cell
		PdfBookmarkEntry bookmarkEntry = new PdfBookmarkEntry();
		bookmarkEntry.setText("Text");
		bookmarkEntry.setDestination(cell);
		bookmarkEntry.setDestinationName("AsposeCells--" + cell.getName());
		  
		//Access cell G56
		cell = ws.getCells().get("G56");
		  
		//Create Sub-Bookmark and Destination for this cell
		PdfBookmarkEntry subbookmarkEntry1 = new PdfBookmarkEntry();
		subbookmarkEntry1.setText("Text1");
		subbookmarkEntry1.setDestination(cell);
		subbookmarkEntry1.setDestinationName("AsposeCells--" + cell.getName());
		  
		//Access cell L4
		cell = ws.getCells().get("L4");
		  
		//Create Sub-Bookmark and Destination for this cell
		PdfBookmarkEntry subbookmarkEntry2 = new PdfBookmarkEntry();
		subbookmarkEntry2.setText("Text2");
		subbookmarkEntry2.setDestination(cell);
		subbookmarkEntry2.setDestinationName("AsposeCells--" + cell.getName());
		  
		//Add Sub-Bookmarks in list
		ArrayList list = new ArrayList();
		list.add(subbookmarkEntry1);
		list.add(subbookmarkEntry2);
		  
		//Assign Sub-Bookmarks list to Bookmark Sub-Entry
		bookmarkEntry.setSubEntry(list);
		  
		//Create PdfSaveOptions and assign Bookmark to it
		PdfSaveOptions opts = new PdfSaveOptions();
		opts.setBookmark(bookmarkEntry);
		  
		//Save the workbook in Pdf format with given pdf save options
		wb.save(outDir + "outputPdfBookmarkEntry_DestinationName.pdf", opts);

		// Print the message
		System.out.println("AddPDFBookmarksWithNamedDestinations executed successfully.");
	}
}
