package AsposeCellsExamples.Worksheets;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;

import javax.imageio.ImageIO;

import com.aspose.cells.ODSPageBackground;
import com.aspose.cells.ODSPageBackgroundGraphicPositionType;
import com.aspose.cells.ODSPageBackgroundType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

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

        ODSPageBackground background = worksheet.getPageSetup().getODSPageBackground();

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
		if(type == ODSPageBackgroundType.COLOR) {
			value = "COLOR";
		} else if(type == ODSPageBackgroundType.GRAPHIC) {
			value = "GRAPHIC";
		} else if(type == ODSPageBackgroundType.NONE) {
			value = "NONE";
		}
		return value;
	}
	
	public static String getPositionValue(int position) {
		String value = "";
		if(position == ODSPageBackgroundGraphicPositionType.BOTTOM_CENTER) {
			value = "BOTTOM_CENTER";
		} else if(position == ODSPageBackgroundGraphicPositionType.BOTTOM_LEFT) {
			value = "BOTTOM_LEFT";
		} else if(position == ODSPageBackgroundGraphicPositionType.BOTTOM_RIGHT) {
			value = "BOTTOM_RIGHT";
		} else if(position == ODSPageBackgroundGraphicPositionType.CENTER_CENTER) {
			value = "CENTER_CENTER";
		} else if(position == ODSPageBackgroundGraphicPositionType.CENTER_LEFT) {
			value = "CENTER_LEFT";
		} else if(position == ODSPageBackgroundGraphicPositionType.CENTER_RIGHT) {
			value = "CENTER_RIGHT";
		} else if(position == ODSPageBackgroundGraphicPositionType.TOP_CENTER) {
			value = "TOP_CENTER";
		} else if(position == ODSPageBackgroundGraphicPositionType.TOP_LEFT) {
			value = "TOP_LEFT";
		} else if(position == ODSPageBackgroundGraphicPositionType.TOP_RIGHT) {
			value = "TOP_RIGHT";
		}
		return value;
	}
    // ExEnd:1
}
