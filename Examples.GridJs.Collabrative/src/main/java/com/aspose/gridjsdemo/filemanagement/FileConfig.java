package com.aspose.gridjsdemo.filemanagement;

import java.io.File;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.security.crypto.bcrypt.BCryptPasswordEncoder;
import org.springframework.security.crypto.password.PasswordEncoder;

import com.aspose.gridjs.CoWorkUserProvider;
import com.aspose.gridjs.GridJsOptions;
 

 
@Configuration
public class FileConfig {
	@Value("${fileconfig.CachePath}")
	public String cachePath;

    
	@Value("${fileconfig.AsposeLicensePath}")
	public String asposeLicensePath;
	//when use jar it shall be added
 
 
 
	@Bean
    public GridJsOptions gridJsOptions() {
		
		com.aspose.cells.License lic = new com.aspose.cells.License();
		if ((new File(asposeLicensePath)).exists()) {
			lic.setLicense(asposeLicensePath);
		}
    	GridJsOptions options=new GridJsOptions();
    	options.setFileCacheDirectory(cachePath);
    	
    	//do not forget here we use Collaborative mode
    	options.setCollaborative(true);
    	options.setLazyLoading(false);
        return options;
    }
 
    
    @Bean
    public CoWorkUserProvider currentUserProvider() {
        return new MyCustomUser();
    }
    
    @Bean
    public PasswordEncoder passwordEncoder() {
        return new BCryptPasswordEncoder();
    }
    
 
}
