package com.aspose.gridjs.demo;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import com.aspose.gridjs.GridJsOptions;

@Configuration
public class MyConfig {
	@Value("${testconfig.CachePath}")
	public String cachePath;

    
	@Value("${testconfig.AsposeLicensePath}")
	public String asposeLicensePath;
	
	
	@Bean
    public GridJsOptions gridJsOptions() {
    	GridJsOptions options=new GridJsOptions();
    	options.setFileCacheDirectory(cachePath);
        return options;
    }


 
}
