package AsposeCellsExamples.DrawingObjects;

import com.aspose.cells.Color;
import com.aspose.cells.FillFormat;
import com.aspose.cells.GradientStyleType;
import com.aspose.cells.LineFormat;
import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.MsoLineDashStyle;
import com.aspose.cells.MsoLineStyle;
import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.PlacementType;
import com.aspose.cells.TextBox;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class ReplaceTagWithTextInTextBox {
    static String srcDir = Utils.Get_SourceDirectory();
    static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {
		// ExStart:AddingTextBoxControl
        Workbook wb = new Workbook(srcDir + "sampleReplaceTagWithText.xlsx");
        String tag = "TAG_2#TAG_1";
        String replace = "1#ys";
        String[] tags = tag.split("#");
        String tag1 = tags[0];
        String tag2 = tags[1];
        for (int i = 0; i < tag.split("#").length; i++)
        {
            sheetReplace(wb, "<" + tag.split("#")[i] + ">", replace.split("#")[i]);
        }
        PdfSaveOptions opts = new PdfSaveOptions();

        wb.save(outDir + "outputReplaceTagWithText.pdf", opts);
		// ExEnd:AddingTextBoxControl
        
        // Print the message
		System.out.println("ReplaceTagWithTextInTextBox executed successfully.");

	}
    public static void sheetReplace(Workbook workbook, String sFind, String sReplace) throws Exception
    {
        String finding = sFind;

        for (Object obj : workbook.getWorksheets())
        {
        	Worksheet sheet = (Worksheet)obj;
            sheet.replace(finding, sReplace);

            for (int j = 0; j < 3; j++)
            {

                if (sheet.getPageSetup().getHeader(j) != null)

                    sheet.getPageSetup().setHeader(j, sheet.getPageSetup().getHeader(j).replace(finding, sReplace));

                if (sheet.getPageSetup().getFooter(j) != null)

                    sheet.getPageSetup().setFooter(j, sheet.getPageSetup().getFooter(j).replace(finding, sReplace));

            }

        }

        for (Object obj: workbook.getWorksheets())
        {
        	Worksheet sheet = (Worksheet)obj;	
            sFind = sFind.replace("<", "&lt;");
            sFind = sFind.replace(">", "&gt;");

            for (Object obj1 : sheet.getTextBoxes())
            {
            	TextBox mytextbox = (TextBox)obj1;

                if (mytextbox.getHtmlText() != null)
                {
                    if (mytextbox.getHtmlText().indexOf(sFind) >= 0)
                    {
                        mytextbox.setHtmlText(mytextbox.getHtmlText().replace(sFind, sReplace));
                    }
                }
            }
        }
    }

}
