package com.aspose.cells.examples.featurescomparison.rowscolumns.addcomments;

import com.aspose.cells.Comment;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AsposeAddCommentsToCell 
{
    public static void main(String[] args) throws Exception 
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeAddCommentsToCell.class);

        //Instantiating a Workbook object
        Workbook workbook = new Workbook();

        Worksheet worksheet = workbook.getWorksheets().get(0);

        //Adding a comment to "F5" cell
        int commentIndex = worksheet.getComments().add("F5");
        Comment comment = worksheet.getComments().get(commentIndex);

        //Setting the comment note
        comment.setNote("Hello Aspose!");

        //Saving the Excel file
        workbook.save(dataDir + "AsposeComments.xls");

        System.out.println("Done.");
    }
}
