package AsposeCellsExamples.Worksheets;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;

import javax.imageio.ImageIO;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;
public class SetODSGraphicBackground {
	
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the source directory.
		String sourceDir = Utils.Get_SourceDirectory();
		// The path to the output directory.
		String outDir = Utils.Get_OutputDirectory();
		
		// Instantiating a Workbook object
        Workbook workbook = new Workbook();

        //Access first worksheet
        Worksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getCells().get(0, 0).setValue(1);
        worksheet.getCells().get(1, 0).setValue(2);
        worksheet.getCells().get(2, 0).setValue(3);
        worksheet.getCells().get(3, 0).setValue(4);
        worksheet.getCells().get(4, 0).setValue(5);
        worksheet.getCells().get(5, 0).setValue(6);
        worksheet.getCells().get(0, 1).setValue(7);
        worksheet.getCells().get(1, 1).setValue(8);
        worksheet.getCells().get(2, 1).setValue(9);
        worksheet.getCells().get(3, 1).setValue(10);
        worksheet.getCells().get(4, 1).setValue(11);
        worksheet.getCells().get(5, 1).setValue(12);

        OdsPageBackground background = worksheet.getPageSetup().getODSPageBackground();
        
        BufferedImage image = ImageIO.read(new File(sourceDir + "background.png"));
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        ImageIO.write(image, "png", bos );
        byte [] imageData = bos.toByteArray();

        background.setType(OdsPageBackgroundType.GRAPHIC);
        background.setGraphicData(imageData);
        background.setGraphicType(OdsPageBackgroundGraphicType.AREA);

        workbook.save(outDir + "GraphicBackground.ods", SaveFormat.ODS);
        // ExEnd:1

        System.out.println("SetODSColoredBackground executed successfully.");
	}
}
