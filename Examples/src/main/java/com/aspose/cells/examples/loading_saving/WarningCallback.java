package com.aspose.cells.examples.loading_saving;

import com.aspose.cells.IWarningCallback;
import com.aspose.cells.WarningInfo;
import com.aspose.cells.WarningType;

//Implement IWarningCallback interface to catch warnings while loading workbook
public class WarningCallback implements IWarningCallback
{
public void warning(WarningInfo warningInfo)
{
    if(warningInfo.getWarningType() == WarningType.DUPLICATE_DEFINED_NAME)
    {
        System.out.println("Duplicate Defined Name Warning: " + warningInfo.getDescription());
    }            
}
}//WarningCallback

