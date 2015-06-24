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

public class WholeNumberDataValidation {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(WholeNumberDataValidation.class);

        //Instantiating an Workbook object
        Workbook workbook = new Workbook();
        WorksheetCollection worksheets = workbook.getWorksheets();

        //Accessing the Validations collection of the worksheet
        Worksheet worksheet = worksheets.get(0);
        ValidationCollection validations = worksheet.getValidations();

        //Creating a Validation object
        int index = validations.add();
        Validation validation = validations.get(index);

        //Setting the validation type to whole number
        validation.setType(ValidationType.WHOLE_NUMBER);

        //Setting the operator for validation to Between
        validation.setOperator(OperatorType.BETWEEN);

        //Setting the minimum value for the validation
        validation.setFormula1("10");

        //Setting the maximum value for the validation
        validation.setFormula2("1000");

        //Applying the validation to a range of cells from A1 to B2 using the
        //CellArea structure
        CellArea area = new CellArea();
        area.StartRow = 0;
        area.StartColumn = 0;
        area.EndRow = 1;
        area.EndColumn = 1;

        //Adding the cell area to Validation
        validation.addArea(area);

        //Saving the Excel file
        workbook.save(dataDir + "output.xls");

        // Print message
        System.out.println("Process completed successfully");

    }
}
