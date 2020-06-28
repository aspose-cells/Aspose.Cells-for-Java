package AsposeCellsExamples.TechnicalArticles;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

import java.io.FileOutputStream;

public class GenerateDatabarImage {
	public static void main(String[] args) throws Exception {

		// ExStart:1
		String sourceDir = Utils.Get_SourceDirectory();
		String outputDir = Utils.Get_OutputDirectory();

		Workbook workbook = new Workbook(sourceDir + "sampleGenerateDatabarImage.xlsx");

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access the cell which contains conditional formatting databar
		Cell cell = worksheet.getCells().get("C1");

		// Create and get the conditional formatting of the worksheet
		int idx = worksheet.getConditionalFormattings().add();
		FormatConditionCollection fcc = worksheet.getConditionalFormattings().get(idx);
		fcc.addCondition(FormatConditionType.DATA_BAR);
		fcc.addArea(CellArea.createCellArea("C1", "C4"));

		// Access the conditional formatting databar
		DataBar dbar = fcc.get(0).getDataBar();

		// Create image or print options
		ImageOrPrintOptions opts = new ImageOrPrintOptions();
		opts.setImageType(ImageType.PNG);

		// Get the image bytes of the databar
		byte[] imgBytes = dbar.toImage(cell, opts);

		// Write image bytes on the disk
		FileOutputStream out = new FileOutputStream(outputDir + "databar.png");
		out.write(imgBytes);
		out.close();

		// save workbook with databars
		workbook.save(outputDir + "databar.xlsx");
		// ExEnd:1

		// Print the message
		System.out.println("GenerateDatabarImage executed successfully.");

	}
}
