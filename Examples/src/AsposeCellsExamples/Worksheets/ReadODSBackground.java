package AsposeCellsExamples.Worksheets;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;

import javax.imageio.ImageIO;

import com.aspose.cells.*;

import AsposeCellsExamples.Utils;
public class ReadODSBackground {

	// ExStart:1
	public static void main(String[] args) throws Exception {
		// The path to the source directory.
		String sourceDir = Utils.Get_SourceDirectory();
		// The path to the output directory.
		String outDir = Utils.Get_OutputDirectory();

        //Load source Excel file
        Workbook workbook = new Workbook(sourceDir + "GraphicBackground.ods");

        //Access first worksheet
        Worksheet worksheet = workbook.getWorksheets().get(0);

        OdsPageBackground background = worksheet.getPageSetup().getODSPageBackground();

        System.out.println("Background Type: " + getTypeValue(background.getType()));
        System.out.println("Backgorund Position: " + getPositionValue(background.getGraphicPositionType()));

        //Save background image
        
        ByteArrayInputStream stream = new ByteArrayInputStream(background.getGraphicData());
        BufferedImage image = ImageIO.read(stream);
        ImageIO.write(image, "png", new File(outDir + "background.png"));

        System.out.println("ReadODSBackground executed successfully.");        
	}
	
	public static String getTypeValue(int type) {
		String value = "";
		if(type == OdsPageBackgroundType.COLOR) {
			value = "COLOR";
		} else if(type == OdsPageBackgroundType.GRAPHIC) {
			value = "GRAPHIC";
		} else if(type == OdsPageBackgroundType.NONE) {
			value = "NONE";
		}
		return value;
	}
	
	public static String getPositionValue(int position) {
		String value = "";
		if(position == OdsPageBackgroundGraphicPositionType.BOTTOM_CENTER) {
			value = "BOTTOM_CENTER";
		} else if(position == OdsPageBackgroundGraphicPositionType.BOTTOM_LEFT) {
			value = "BOTTOM_LEFT";
		} else if(position == OdsPageBackgroundGraphicPositionType.BOTTOM_RIGHT) {
			value = "BOTTOM_RIGHT";
		} else if(position == OdsPageBackgroundGraphicPositionType.CENTER_CENTER) {
			value = "CENTER_CENTER";
		} else if(position == OdsPageBackgroundGraphicPositionType.CENTER_LEFT) {
			value = "CENTER_LEFT";
		} else if(position == OdsPageBackgroundGraphicPositionType.CENTER_RIGHT) {
			value = "CENTER_RIGHT";
		} else if(position == OdsPageBackgroundGraphicPositionType.TOP_CENTER) {
			value = "TOP_CENTER";
		} else if(position == OdsPageBackgroundGraphicPositionType.TOP_LEFT) {
			value = "TOP_LEFT";
		} else if(position == OdsPageBackgroundGraphicPositionType.TOP_RIGHT) {
			value = "TOP_RIGHT";
		}
		return value;
	}
    // ExEnd:1
}
