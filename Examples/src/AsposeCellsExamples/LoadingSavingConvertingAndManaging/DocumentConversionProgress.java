package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

public class DocumentConversionProgress {
	public static void main(String[] args) throws Exception 
	{
		// ExStart:1
		// The path to the source directory.
		String sourceDir = Utils.Get_SourceDirectory();

		// The path to the output directory.
		String outputDir = Utils.Get_OutputDirectory();

		Workbook wb = new Workbook(sourceDir + "PagesBook1.xlsx");

		PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
		pdfSaveOptions.setPageSavingCallback(new TestPageSavingCallback());

		wb.save(outputDir + "DocumentConversionProgress.pdf", pdfSaveOptions);
		// ExEnd:1
	}
}

// ExStart:2
class TestPageSavingCallback implements IPageSavingCallback {
    public void pageStartSaving(PageStartSavingArgs args)
    {
        System.out.println("Start saving page index " + args.getPageIndex() + " of pages " + args.getPageCount());

        //don't output pages before page index 2.
        if (args.getPageIndex() < 2)
        {
            args.setToOutput(false);
        }
    }

    public void pageEndSaving(PageEndSavingArgs args)
    {
		System.out.println("End saving page index " + args.getPageIndex() + " of pages " + args.getPageCount());

        //don't output pages after page index 8.
        if (args.getPageIndex() >= 8)
        {
            args.setHasMorePages(false);
        }
    }
}
// ExEnd:2