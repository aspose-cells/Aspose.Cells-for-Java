package AsposeCellsExamples.DrawingObjects;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ReplaceTextInSmartArt {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

        // ExStart:1
        Workbook wb = new Workbook(srcDir + "SmartArt.xlsx");
        for (Object obj : wb.getWorksheets())
        {
        	Worksheet worksheet = (Worksheet)obj;
            for (Object shp : worksheet.getShapes())
            {
            	Shape shape = (Shape)shp;
                shape.setAlternativeText("ReplacedAlternativeText"); // This works fine just as the normal Shape objects do.
                if (shape.isSmartArt())
                {
                    for (Shape smartart : shape.getResultOfSmartArt().getGroupedShapes())
                    {
                        smartart.setText("ReplacedText"); // This doesn't update the text in Workbook which I save to the another file.
                    }
                }
            }
        }
        com.aspose.cells.OoxmlSaveOptions options = new com.aspose.cells.OoxmlSaveOptions();
        options.setUpdateSmartArt(true);
        
        wb.save(outDir + "outputSmartArt.xlsx", options);
        // ExEnd:1
        
     	// Print message
     	System.out.println("ReplaceTextInSmartArt executed successfully");        

	}
}
