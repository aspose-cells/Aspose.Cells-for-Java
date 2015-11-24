package com.aspose.cells.examples.featurescomparison.formatting.colors;

import com.aspose.cells.BackgroundType;
import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.Font;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AsposeWorkingWithColors
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeWorkingWithColors.class);

        //Instantiating a Workbook object
        Workbook workbook = new Workbook();

        //Accessing the added worksheet in the Excel file
        Worksheet worksheet = workbook.getWorksheets().get(0);
        Cells cells = worksheet.getCells();

        // === Setting Background Pattern ===

        //Accessing cell from the worksheet
        Cell cell = cells.get("B2");
        Style style = cell.getStyle();

        //Setting the foreground color to yellow
        style.setBackgroundColor(Color.getYellow());

        //Setting the background pattern to vertical stripe
        style.setPattern(BackgroundType.VERTICAL_STRIPE);

        //Saving the modified style to the "A1" cell.
        cell.setStyle(style);

        // === Setting Foreground ===

        //Adding custom color to the palette at 55th index
        Color color = Color.fromArgb(212,213,0);
        workbook.changePalette(color,55);

        //Accessing the "A2" cell from the worksheet
        cell = cells.get("B3");

        //Adding some value to the cell
        cell.setValue("Hello Aspose!");

        //Setting the custom color to the font
        style = cell.getStyle();
        Font font = style.getFont();
        font.setColor(color);

        cell.setStyle(style);

        //Saving the Excel file
        workbook.save(dataDir + "AsposeColors_Out.xls");

        System.out.println("Aspose Colors Created.");
    }
}
