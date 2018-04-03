package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class SpecifyDocumentVersionOfExcelFile {
	
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
	
		//Create workbook object
		Workbook wb = new Workbook();

		//Access built-in document property collection
		BuiltInDocumentPropertyCollection bdpc = wb.getBuiltInDocumentProperties();

		//Set the title
		bdpc.setTitle("Aspose File Format APIs");

		//Set the author
		bdpc.setAuthor("Aspose APIs Developers");

		//Set the document version
		bdpc.setDocumentVersion("Aspose.Cells Version - 18.3");

		//Save the workbook in xlsx format
		wb.save(outDir + "outputSpecifyDocumentVersionOfExcelFile.xlsx", SaveFormat.XLSX);

		// Print the message
		System.out.println("SpecifyDocumentVersionOfExcelFile executed successfully.");
	}
}
