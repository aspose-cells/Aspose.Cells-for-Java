package AsposeCellsExamples.Data;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ChangeCellsAlignmentAndKeepExistingFormatting { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		// Load sample Excel file containing cells with formatting.
		Workbook wb = new Workbook(srcDir + "sampleChangeCellsAlignmentAndKeepExistingFormatting.xlsx");

		// Access first worksheet.
		Worksheet ws = wb.getWorksheets().get(0);

		// Create cells range.
		Range rng = ws.getCells().createRange("B2:D7");

		// Create style object.
		Style st = wb.createStyle();

		// Set the horizontal and vertical alignment to center.
		st.setHorizontalAlignment(TextAlignmentType.CENTER);
		st.setVerticalAlignment(TextAlignmentType.CENTER);

		// Create style flag object.
		StyleFlag flag = new StyleFlag();

		// Set style flag alignments true. It is most crucial statement.
		// Because if it will be false, no changes will take place.
		flag.setAlignments(true);

		// Apply style to range of cells.
		rng.applyStyle(st, flag);

		// Save the workbook in XLSX format.
		wb.save(outDir + "outputChangeCellsAlignmentAndKeepExistingFormatting.xlsx", SaveFormat.XLSX);
	
		// Print the message
		System.out.println("ChangeCellsAlignmentAndKeepExistingFormatting executed successfully.");
	}
}
