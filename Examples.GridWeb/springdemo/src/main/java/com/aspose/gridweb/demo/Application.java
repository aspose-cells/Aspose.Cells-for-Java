package com.aspose.gridweb.demo;
 
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.boot.web.servlet.ServletRegistrationBean;
import org.springframework.context.ApplicationContext;
import org.springframework.context.annotation.Bean;
import org.springframework.stereotype.Component;

import com.aspose.gridweb.ExtPage;
//import com.aspose.gridweb.GridWebServlet;
//import com.aspose.gridweb.test.servlet.FeatureServlet;
//import com.aspose.gridweb.test.servlet.SheetsServlet;
import com.aspose.gridweb.GridWebServlet;
import com.aspose.gridweb.ManualLog;
import com.aspose.gridweb.test.servlet.FeatureServlet;
import com.aspose.gridweb.test.servlet.FormatServlet;
import com.aspose.gridweb.test.servlet.FunctionServlet;
import com.aspose.gridweb.test.servlet.SheetsServlet;
import com.aspose.gridweb.test.servlet.WebCellsServlet;
@Component
@ConfigurationProperties(prefix = "testconfig")
class TestConfig {
    private String logPath;
    private String cachePath;

    public String getLogPath() {
        return logPath;
    }

    public void setLogPath(String logPath) {
        this.logPath = logPath;
    }

    public String getCachePath() {
        return cachePath;
    }

    public void setCachePath(String cachePath) {
        this.cachePath = cachePath;
    }
}
@SpringBootApplication
public class Application {
 
	    @Bean
	    public ServletRegistrationBean servletRegistrationBean() {
	        return new ServletRegistrationBean(new GridWebServlet(), "/gridwebdemo/GridWebServlet/*");
	    }
	    @Bean
	    public ServletRegistrationBean servletRegistrationBean2() {
	    	return new ServletRegistrationBean(new SheetsServlet(), "/gridwebdemo/SheetsServlet/*");
	    }
	    @Bean
	    public ServletRegistrationBean servletRegistrationBean3() {
	    	return new ServletRegistrationBean(new FeatureServlet(), "/gridwebdemo/FeatureServlet/*");
	    }
	    @Bean
	    public ServletRegistrationBean servletRegistrationBean4() {
	    	return new ServletRegistrationBean(new WebCellsServlet(), "/gridwebdemo/WebCellsServlet/*");
	    }
	    @Bean
	    public ServletRegistrationBean servletRegistrationBean5() {
	    	return new ServletRegistrationBean(new FunctionServlet(), "/gridwebdemo/FunctionServlet/*");
	    }
	    
	    @Bean
	    public ServletRegistrationBean servletRegistrationBean6() {
	    	return new ServletRegistrationBean(new FormatServlet(), "/gridwebdemo/FormatServlet/*");
	    }
	    
	    
    public static void main(String[] args) {
    	
    	ApplicationContext context =  SpringApplication.run(Application.class, args);
    	 
		// com.aspose.gridweb.License li=new com.aspose.gridweb.License();
		// li.setLicense("D:\\release\\Aspose.Total.Product.Family.lic");
		//optional settings for cache
    	  TestConfig config = context.getBean(TestConfig.class);

          System.out.println("Log Path: " + config.getLogPath());
          System.out.println("Cache Path: " + config.getCachePath());
          
		ExtPage.setMaxholders(1000);
		ExtPage.setMemoryInstanceExpires(600l);
		ExtPage.setMemoryCleanRateTime(1200l);
		//#######the  dir for cache store for spreadsheet files,make sure the directory is existed at you enviroment.#############
		ExtPage.setTempfilepath(config.getCachePath());
        //set log directory, optional 
        ManualLog.setBasicPathAndInit(config.getLogPath());
       
    }
    
 
}