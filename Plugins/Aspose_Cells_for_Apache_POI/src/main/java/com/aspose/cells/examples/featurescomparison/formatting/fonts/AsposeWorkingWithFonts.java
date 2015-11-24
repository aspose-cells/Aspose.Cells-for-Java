package com.aspose.cells.examples.featurescomparison.formatting.fonts;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.Font;
import com.aspose.cells.FontUnderlineType;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AsposeWorkingWithFonts
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeWorkingWithFonts.class);

        //Instantiating a Workbook object
        Workbook workbook = new Workbook();

        //Accessing the worksheet in the Excel file
        Worksheet worksheet = workbook.getWorksheets().get(0);
        Cells cells = worksheet.getCells();

        //Adding some value to cell
        Cell cell = cells.get("A1");
        cell.setValue("This is Aspose test of fonts!");

        //Setting the font name to "Times New Roman"
        Style style = cell.getStyle();
        Font font = style.getFont();
        font.setName("Courier New");
        font.setSize(24);
        font.setBold(true);
        font.setUnderline(FontUnderlineType.SINGLE);
        font.setColor(Color.getBlue());
        font.setStrikeout(true);
        //font.setSubscript(true);

        cell.setStyle(style); 

        //Saving the modified Excel file in default format
        workbook.save(dataDir + "AsposeFonts.xls");

        System.out.println("Aspose Fonts Created.");
    }
}
