package AsposeCellsExamples.TechnicalArticles;

import java.util.Locale;

import com.aspose.cells.ChartGlobalizationSettings;

public class CustomChartGlobalizationSettings extends ChartGlobalizationSettings
{
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
