package AsposeCellsExamples.Worksheets.PageSetupFeatures;

import java.util.*;
import com.aspose.cells.*;

public class CopyPageSetupSettingsFromSourceWorksheetToDestinationWorksheet {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		//Converting integer enums to string enums
		HashMap<Integer, String> paperSizeTypes = new HashMap<>();
		paperSizeTypes.put(PaperSizeType.PAPER_A_3_EXTRA_TRANSVERSE, "PAPER_A_3_EXTRA_TRANSVERSE");
		paperSizeTypes.put(PaperSizeType.PAPER_LETTER, "PAPER_LETTER");
		 
		//Create workbook
		Workbook wb = new Workbook();
		 
		//Add two test worksheets
		wb.getWorksheets().add("TestSheet1");
		wb.getWorksheets().add("TestSheet2");
		 
		//Access both worksheets as TestSheet1 and TestSheet2
		Worksheet TestSheet1 = wb.getWorksheets().get("TestSheet1");
		Worksheet TestSheet2 = wb.getWorksheets().get("TestSheet2");
		 
		//Set the Paper Size of TestSheet1 to PaperA3ExtraTransverse
		TestSheet1.getPageSetup().setPaperSize(PaperSizeType.PAPER_A_3_EXTRA_TRANSVERSE);
		 
		//Print the Paper Size of both worksheets
		System.out.println("Before Paper Size: " + paperSizeTypes.get(TestSheet1.getPageSetup().getPaperSize()));
		System.out.println("Before Paper Size: " + paperSizeTypes.get(TestSheet2.getPageSetup().getPaperSize()));
		System.out.println();
		 
		//Copy the PageSetup from TestSheet1 to TestSheet2
		TestSheet2.getPageSetup().copy(TestSheet1.getPageSetup(), new CopyOptions());
		 
		//Print the Paper Size of both worksheets
		System.out.println("After Paper Size: " + paperSizeTypes.get(TestSheet1.getPageSetup().getPaperSize()));
		System.out.println("After Paper Size: " + paperSizeTypes.get(TestSheet2.getPageSetup().getPaperSize()));
		System.out.println();
		// ExEnd:1
		
		//Print the message
		System.out.println("CopyPageSetupSettingsFromSourceWorksheetToDestinationWorksheet executed successfully.");
	}
}
