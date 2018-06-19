package AsposeCellsExamples.Fonts;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class SpecifyIndividualOrPrivateSetOfFontsForWorkbookRendering { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		// Path of your custom font directory.
		String customFontsDir = srcDir + "CustomFonts";
		 
		// Specify individual font configs custom font directory.
		IndividualFontConfigs fontConfigs = new IndividualFontConfigs();
		 
		// If you comment this line or if custom font directory is wrong or 
		// if it does not contain required font then output pdf will not be rendered correctly.
		fontConfigs.setFontFolder(customFontsDir, false);
		 
		// Specify load options with font configs.
		LoadOptions opts = new LoadOptions(LoadFormat.XLSX);
		opts.setFontConfigs(fontConfigs);
		 
		// Load the sample Excel file with individual font configs. 
		Workbook wb = new Workbook(srcDir + "sampleSpecifyIndividualOrPrivateSetOfFontsForWorkbookRendering.xlsx", opts);
		 
		// Save to pdf format.
		wb.save(outDir + "outputSpecifyIndividualOrPrivateSetOfFontsForWorkbookRendering.pdf", SaveFormat.PDF);
		
		// Print the message
		System.out.println("SpecifyIndividualOrPrivateSetOfFontsForWorkbookRendering executed successfully.");
	}
}
