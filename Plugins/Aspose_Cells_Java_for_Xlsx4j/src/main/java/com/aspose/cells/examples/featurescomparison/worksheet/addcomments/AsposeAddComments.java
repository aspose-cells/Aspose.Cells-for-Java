package com.aspose.cells.examples.featurescomparison.worksheet.addcomments;

import com.aspose.cells.Comment;
import com.aspose.cells.Font;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AsposeAddComments
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeAddComments.class);

        //Instantiating a Workbook object
        Workbook workbook = new Workbook();

        //Adding a new worksheet to the Workbook object
        Worksheet worksheet = workbook.getWorksheets().get(0);

        //Adding a comment to cell
        int commentIndex = worksheet.getComments().add("A1");
        Comment comment = worksheet.getComments().get(commentIndex);

        //Setting the comment note
        comment.setNote("Hello Aspose!");

        //Setting the font size of a comment to 14
        Font font = comment.getFont();
        font.setSize(14);
        //Setting the font of a comment to bold
        font.setBold(true);

        //Setting the height of the font to 10
        comment.setHeightCM(10);

        //Setting the width of the font to 2
        comment.setWidthCM(2);

        //Saving the Excel file
        workbook.save(dataDir + "AddComments-Aspose.xlsx", SaveFormat.XLSX);

        //Print Message
        System.out.println("Comment added successfully.");
    }
}
