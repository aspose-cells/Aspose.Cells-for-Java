package com.aspose.gridweb.demo;

import org.springframework.context.annotation.Configuration;
import org.springframework.web.servlet.config.annotation.ResourceHandlerRegistry;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurationSupport;
 
 

/**
 * @author cailei.lu
 * @description
 * @date 2018/8/3
 */
@Configuration
public class WebAppConfiguration extends WebMvcConfigurationSupport {

    @Override
    public void addResourceHandlers(ResourceHandlerRegistry registry) {
    	 registry.addResourceHandler(new String[] { "/gridwebdemo/**" }).addResourceLocations(new String[] { "classpath:/static/" });
    	 super.addResourceHandlers(registry);
    }


}
 
