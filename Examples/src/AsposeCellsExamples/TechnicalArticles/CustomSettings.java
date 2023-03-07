package AsposeCellsExamples.TechnicalArticles;

import java.util.Locale;

import com.aspose.cells.ConsolidationFunction;
import com.aspose.cells.GlobalizationSettings;

public class CustomSettings extends GlobalizationSettings {
	public String getTotalName(int functionType) {
		switch (functionType) {
		case ConsolidationFunction.AVERAGE:
			return "AVG";

		// Handle other cases

		default:
			return super.getTotalName(functionType);
		}
	}

	public String getGrandTotalName(int functionType) {
		switch (functionType) {
		case ConsolidationFunction.AVERAGE:
			return "GRAND AVG";

		// Handle other cases

		default:
			return super.getGrandTotalName(functionType);
		}
	}
	
	public String getOtherName()
    {
        String language = Locale.getDefault().getLanguage();
		System.out.println(language);
		switch (language)
		{
		    case "en":
		        return "Other";
		    case "fr":
		        return "Autre";
		    case "de":
		        return "Andere";
		
		    //Handle other cases as per requirement
		
		    default:
					return "Other"; //default to English
		}
    }
}
