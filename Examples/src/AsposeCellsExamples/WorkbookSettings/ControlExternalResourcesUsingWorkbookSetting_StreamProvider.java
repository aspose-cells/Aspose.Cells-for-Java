package AsposeCellsExamples.WorkbookSettings;

import java.io.*;
import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ControlExternalResourcesUsingWorkbookSetting_StreamProvider { 
	
	String srcDir = Utils.Get_SourceDirectory();
	String outDir = Utils.Get_OutputDirectory();
	
	//Implementation of IStreamProvider
	class SP implements IStreamProvider
	{
		public void closeStream(StreamProviderOptions arg0) throws Exception {		
		}

		public void initStream(StreamProviderOptions options) throws Exception {

			//Open the filestream of Aspose Logo and assign it to StreamProviderOptions.Stream property
			File imgFile = new File(srcDir + "sampleControlExternalResourcesUsingWorkbookSetting_StreamProvider.png");

			byte[] bts = new byte[(int) imgFile.length()];

			FileInputStream fin = new FileInputStream(imgFile);
			fin.read(bts);
			fin.close();

			ByteArrayOutputStream baout = new ByteArrayOutputStream();
			baout.write(bts);
			baout.close();
			
			options.setStream(baout);
		}
	}


	public void Run() throws Exception {
		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());

		//String srcDir = Utils.Get_SourceDirectory();
		//String outDir = Utils.Get_OutputDirectory();

		//Load sample Excel file containing the external resource e.g. linked image etc.
		Workbook wb = new Workbook(srcDir + "sampleControlExternalResourcesUsingWorkbookSetting_StreamProvider.xlsx");

		//Provide your implementation of IStreamProvider
		wb.getSettings().setStreamProvider(new SP());

		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		//Specify image or print options, we need one page per sheet and png output
		ImageOrPrintOptions opts = new ImageOrPrintOptions();
		opts.setOnePagePerSheet(true);
		opts.setImageFormat(ImageFormat.getPng());

		//Create sheet render by passing required parameters
		SheetRender sr = new SheetRender(ws, opts);

		//Convert your entire worksheet into png image
		sr.toImage(0, outDir + "outputControlExternalResourcesUsingWorkbookSettingStreamProvider.png");

		// Print the message
		System.out.println("ControlExternalResourcesUsingWorkbookSetting_StreamProvider executed successfully.");
	}

	public static void main(String[] args) throws Exception {
		new ControlExternalResourcesUsingWorkbookSetting_StreamProvider().Run();
	}
}
