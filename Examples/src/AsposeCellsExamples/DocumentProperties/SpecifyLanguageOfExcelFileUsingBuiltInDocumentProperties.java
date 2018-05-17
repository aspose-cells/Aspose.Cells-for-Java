package AsposeCellsExamples.DocumentProperties;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class SpecifyLanguageOfExcelFileUsingBuiltInDocumentProperties { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Create workbook object.
		Workbook wb = new Workbook();

		//Access built-in document property collection.
		BuiltInDocumentPropertyCollection bdpc = wb.getBuiltInDocumentProperties();

		//Set the language of the Excel file.
		bdpc.setLanguage("German, French");

		//Save the workbook in xlsx format.
		wb.save(outDir + "outputSpecifyLanguageOfExcelFileUsingBuiltInDocumentProperties.xlsx", SaveFormat.XLSX);
		 
		// Print the message
		System.out.println("SpecifyLanguageOfExcelFileUsingBuiltInDocumentProperties executed successfully.");
	}
}
