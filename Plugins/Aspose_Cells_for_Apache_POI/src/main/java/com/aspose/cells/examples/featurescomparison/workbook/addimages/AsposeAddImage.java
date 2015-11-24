package com.aspose.cells.examples.featurescomparison.workbook.addimages;

import com.aspose.cells.PlacementType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AsposeAddImage
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeAddImage.class);

        //Instantiate a new workbook
        Workbook workbook = new Workbook();

        //Get the first worksheet
        Worksheet worksheet = workbook.getWorksheets().get(0);

        //Insert a string value to a cell
        worksheet.getCells().get("C2").setValue("Image");

        //Set the 4th row height
        worksheet.getCells().setRowHeight(3, 150);

        //Set the C column width
        worksheet.getCells().setColumnWidth(2,50);

        //Add a picture to the C4 cell
        int index = worksheet.getPictures().add(3, 2, dataDir + "aspose.jpg");

        //Get the picture object
        com.aspose.cells.Picture pic = worksheet.getPictures().get(index);

        //Set the placement type
        pic.setPlacement(PlacementType.FREE_FLOATING);

        //Add an image hyperlink
        pic.addHyperlink("http://www.aspose.com/");
        com.aspose.cells.Hyperlink hlink = pic.getHyperlink();

        //Specify the screen tip
        hlink.setScreenTip("Click to go to Aspose site");

        //Save the Excel file
        workbook.save(dataDir + "AsposeImage.xlsx", SaveFormat.XLSX);

        System.out.println("Done...");
    }
}