package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.GlobalizationSettings;

public class GlobalizationSettingsImp extends GlobalizationSettings{
	//This function will return the sub total name
    public String getTotalName(int functionType)
    {
        return "Chinese Total - 可能的用法";
    }

    //This function will return the grand total name
    public String getGrandTotalName(int functionType)
    {
        return "Chinese Grand Total - 可能的用法";
    }
}
