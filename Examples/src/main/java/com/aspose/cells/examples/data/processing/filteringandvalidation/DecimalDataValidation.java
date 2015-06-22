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

public class DecimalDataValidation {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DecimalDataValidation.class);

        // Create a workbook object.
        Workbook workbook = new Workbook();

        // Create a worksheet and get the first worksheet.
        Worksheet ExcelWorkSheet = workbook.getWorksheets().get(0);

        // Obtain the existing Validations collection.
        ValidationCollection validations = ExcelWorkSheet.getValidations();

        // Create a validation object adding to the collection list.
        int index = validations.add();
        Validation validation = validations.get(index);

        // Set the validation type.
        validation.setType(ValidationType.DECIMAL);

        // Specify the operator.
        validation.setOperator(OperatorType.BETWEEN);

        // Set the lower and upper limits.
        validation.setFormula1(new Double(Double.MIN_VALUE).toString());
        validation.setFormula2(new Double(Double.MAX_VALUE).toString());

        // Set the error message.
        validation.setErrorMessage("Please enter a valid integer or decimal number");

        // Specify the validation area of cells.
        CellArea area = new CellArea();
        area.StartRow = 0;
        area.StartColumn = 0;
        area.EndRow = 9;
        area.EndColumn = 0;

        // Add the area.
        validation.addArea(area);

        // Save the workbook.
        workbook.save(dataDir + "output.xls");

        // Print message
        System.out.println("Process completed successfully");

    }
}
