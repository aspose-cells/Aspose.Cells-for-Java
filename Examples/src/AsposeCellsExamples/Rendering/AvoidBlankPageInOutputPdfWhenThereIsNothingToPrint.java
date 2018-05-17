package AsposeCellsExamples.Rendering;

import java.io.ByteArrayOutputStream;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class AvoidBlankPageInOutputPdfWhenThereIsNothingToPrint {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Create empty workbook.
		Workbook wb = new Workbook();
		
		//Create Pdf save options.
		PdfSaveOptions opts = new PdfSaveOptions();

		//Default value of OutputBlankPageWhenNothingToPrint is true.
		//Setting false means - Do not output blank page when there is nothing to print.
		opts.setOutputBlankPageWhenNothingToPrint(false);

		//Save workbook to Pdf format, it will throw exception because workbook has nothing to print.
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		
		try
		{
			wb.save(baos, opts);
		}
		catch(Exception ex)
		{
			System.out.println("Exception Message: " + ex.getMessage());
		}		
		 
		// Print the message
		System.out.println("AvoidBlankPageInOutputPdfWhenThereIsNothingToPrint executed successfully.");
	}
}
