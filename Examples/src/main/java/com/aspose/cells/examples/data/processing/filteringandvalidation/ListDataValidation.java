/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.cells.examples.data.processing.filteringandvalidation;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class ListDataValidation {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ListDataValidation.class);

        // Create a workbook object.
        Workbook workbook = new Workbook();

        // Get the first worksheet.
        Worksheet ExcelWorkSheet = workbook.getWorksheets().get(0);

        // Add a new worksheet and access it.
        int sheetIndex = workbook.getWorksheets().add();
        Worksheet worksheet2 = workbook.getWorksheets().get(sheetIndex);

        // Create a range with name in the second worksheet.
        Range range = worksheet2.getCells().createRange(0, 4, 4, 4);
        range.setName("MyRange");

        // Fill different cells with data in the range.
        range.get(0, 0).setValue("Blue");
        range.get(1, 0).setValue("Red");
        range.get(2, 0).setValue("Green");
        range.get(3, 0).setValue("Yellow");

        // Obtain the existing Validations collection.
        ValidationCollection validations = ExcelWorkSheet.getValidations();

        // Create a validation object adding to the collection list.
        int index = validations.add();
        Validation validation = validations.get(index);

        // Set the validation type.
        validation.setType(ValidationType.LIST);

        // Set the in cell drop down.
        validation.setInCellDropDown(true);

        // Set the formula1.
        validation.setFormula1("=MyRange");

        // Enable it to show error.
        validation.setShowError(true);

        // Set the alert type severity level.
        validation.setAlertStyle(ValidationAlertType.STOP);

        // Set the error title.
        validation.setErrorTitle("Error");

        // Set the error message.
        validation.setErrorMessage("Please select a color from the list");

        // Specify the validation area of cells.
        CellArea area = new CellArea();
        area.StartRow = 0;
        area.StartColumn = 0;
        area.EndRow = 4;
        area.EndColumn = 0;

        // Add the Validation area.
        validation.addArea(area);

        // Save the excel file.
        workbook.save(dataDir + "output.xls");

        // Print message
        System.out.println("Process completed successfully");

    }
}
