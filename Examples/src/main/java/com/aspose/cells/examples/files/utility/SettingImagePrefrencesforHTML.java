package com.aspose.cells.examples.files.utility;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class SettingImagePrefrencesforHTML {

    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SettingImagePrefrencesforHTML.class);

        //Instantiate a Workbook object by excel file path
        Workbook workbook = new Workbook(dataDir + "Book1.xlsx");

        //Create an instance of HtmlSaveOptions
       /* HtmlSaveOptions  saveOptions = new HtmlSaveOptions(FileFormatType.HTML);
        saveOptions.getImageOptions()

        //Set the ImageFormat to PNG
        saveOptions.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;

        //Set SmoothingMode to AntiAlias
        saveOptions.ImageOptions.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

        //Set TextRenderingHint to AntiAlias
        saveOptions.ImageOptions.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

        //Save spreadsheet to HTML while passing object of HtmlSaveOptions
        book.Save( dataDir + "output.html", saveOptions);*/
        // Print message
        System.out.println("Set PDF Creation Time successfully.");
        //ExEnd:1
    }
}
