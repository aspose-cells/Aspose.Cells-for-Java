package AsposeCellsExamples.Fonts;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class SetDefaultFontPropertyOfPdfSaveOptionsAndImageOrPrintOptions {

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());

		String srcDir = Utils.Get_SourceDirectory();
		String outDir = Utils.Get_OutputDirectory();

		// Open an Excel file.
		Workbook workbook = new Workbook(
				srcDir + "sampleSetDefaultFontPropertyOfPdfSaveOptionsAndImageOrPrintOptions.xlsx");

		// Rendering to PNG file format while setting the
		// CheckWorkbookDefaultFont attribue to false.
		// So, "Times New Roman" font would be used for any missing (not
		// installed) font in the workbook.
		ImageOrPrintOptions imgOpt = new ImageOrPrintOptions();
		imgOpt.setImageFormat(ImageFormat.getPng());
		imgOpt.setCheckWorkbookDefaultFont(false);
		imgOpt.setDefaultFont("Times New Roman");
		SheetRender sr = new SheetRender(workbook.getWorksheets().get(0), imgOpt);
		sr.toImage(0, outDir + "outputSetDefaultFontProperty_ImagePNG.png");

		// Rendering to TIFF file format while setting the
		// CheckWorkbookDefaultFont attribue to false.
		// So, "Times New Roman" font would be used for any missing (not
		// installed) font in the workbook.
		imgOpt.setImageFormat(ImageFormat.getTiff());
		WorkbookRender wr = new WorkbookRender(workbook, imgOpt);
		wr.toImage(outDir + "outputSetDefaultFontProperty_ImageTIFF.tiff");

		// Rendering to PDF file format while setting the
		// CheckWorkbookDefaultFont attribue to false.
		// So, "Times New Roman" font would be used for any missing (not
		// installed) font in the workbook.
		PdfSaveOptions saveOptions = new PdfSaveOptions();
		saveOptions.setDefaultFont("Times New Roman");
		saveOptions.setCheckWorkbookDefaultFont(false);
		workbook.save(outDir + "outputSetDefaultFontProperty_PDF.pdf", saveOptions);

		// Print the message
		System.out.println("SetDefaultFontPropertyOfPdfSaveOptionsAndImageOrPrintOptions executed successfully.");
	}
}
