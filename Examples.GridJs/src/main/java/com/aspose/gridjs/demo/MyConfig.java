package com.aspose.gridjs.demo;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.ComponentScan;
import org.springframework.context.annotation.Configuration;

import com.aspose.gridjs.GridJsOptions;
import com.aspose.gridjs.GridJsService;
import com.aspose.gridjs.GridJsWorkbook;
import com.aspose.gridjs.IGridJsService;

@Configuration
public class MyConfig {
	@Value("${testconfig.CachePath}")
	public String cachePath;

    
	@Value("${testconfig.AsposeLicensePath}")
	public String asposeLicensePath;
	
	
	@Bean
    public GridJsWorkbook gridJsWorkbook() {
        return new GridJsWorkbook();  
    }
	
	
    @Bean
    public GridJsOptions gridJsOptions() {
    	GridJsOptions options=new GridJsOptions();
    	options.setFileCacheDirectory(cachePath);
        return options;
    }

    @Bean
    public IGridJsService gridJsService(GridJsOptions gridJsOptions) throws Exception {
        return new GridJsService(gridJsOptions);
    }
}
