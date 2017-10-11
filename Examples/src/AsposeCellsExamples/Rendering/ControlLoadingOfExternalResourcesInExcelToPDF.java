package AsposeCellsExamples.Rendering;

import java.io.*;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ControlLoadingOfExternalResourcesInExcelToPDF { 

	String srcDir = Utils.Get_SourceDirectory();
	String outDir = Utils.Get_OutputDirectory();

	// Implement IStreamProvider
	class MyStreamProvider implements IStreamProvider {

		public void closeStream(StreamProviderOptions options) throws Exception {
			System.out.println("-----Close Stream-----");
		}

		public void initStream(StreamProviderOptions options) throws Exception {
			System.out.println("-----Init Stream-----");

			// Read the new image in a memory stream and assign it to Stream property
			File imgFile = new File( srcDir + "newPdfSaveOptions_StreamProvider.png");

			byte[] bts = new byte[(int) imgFile.length()];

			FileInputStream fin = new FileInputStream(imgFile);
			fin.read(bts);
			fin.close();

			ByteArrayOutputStream baout = new ByteArrayOutputStream();
			baout.write(bts);
			baout.close();
			
			options.setStream(baout);
		}
	}//MyStreamProvider

	// ------------------------------------------------
	// ------------------------------------------------

	void Run() throws Exception {
		
		// Load source Excel file containing external image
		Workbook wb = new Workbook(srcDir + "samplePdfSaveOptions_StreamProvider.xlsx");

		// Specify Pdf Save Options - Stream Provider
		PdfSaveOptions opts = new PdfSaveOptions();
		opts.setOnePagePerSheet(true);
		opts.setStreamProvider(new MyStreamProvider());

		// Save the workbook to Pdf
		wb.save(outDir + "outputPdfSaveOptions_StreamProvider.pdf", opts);
	}

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());

		ControlLoadingOfExternalResourcesInExcelToPDF pg = new ControlLoadingOfExternalResourcesInExcelToPDF();
		pg.Run();

		// Print the message
		System.out.println("ControlLoadingOfExternalResourcesInExcelToPDF executed successfully.");
	}
}
