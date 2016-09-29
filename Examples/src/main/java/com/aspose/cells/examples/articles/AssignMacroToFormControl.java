package com.aspose.cells.examples.articles;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class AssignMacroToFormControl {

    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getSharedDataDir(AssignMacroToFormControl.class) + "articles/";

        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.getWorksheets().get(0);

        int moduleIdx = workbook.getVbaProject().getModules().add(sheet);
        VbaModule module = workbook.getVbaProject().getModules().get(moduleIdx);
        module.setCodes("Sub ShowMessage()" + "\r\n" +
                "    MsgBox \"Welcome to Aspose!\"" + "\r\n" +
                "End Sub");

        Button button = (Button) sheet.getShapes().addShape(MsoDrawingType.BUTTON, 2, 0, 2, 0, 28, 80);
        button.setPlacement(PlacementType.FREE_FLOATING);
        button.getFont().setName("Tahoma");
        button.getFont().setBold(true);
        button.getFont().setColor(Color.getBlue());
        button.setText("Aspose");

        button.setMacroName(sheet.getName() + ".ShowMessage");

        workbook.save(dataDir + "AMToFControl_out.xlsm");

        System.out.println("File saved");

	}
}
