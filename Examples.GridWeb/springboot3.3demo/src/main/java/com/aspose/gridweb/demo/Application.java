package com.aspose.gridweb.demo;
 
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.web.servlet.ServletRegistrationBean;
import org.springframework.context.annotation.Bean;

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
//import com.aspose.gridweb.ManualLog;
 
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
		 com.aspose.gridweb.License li=new com.aspose.gridweb.License();
//		 li.setLicense("D:\\release\\Aspose.Total.Product.Family.lic");
		 //hello jdk 11
		ExtPage.setMaxholders(1000);
		ExtPage.setMemoryInstanceExpires(600);
		ExtPage.setMemoryCleanRateTime(1200);
		ExtPage.setTempfilepath("D:\\tmpdel\\gridwebjavatmp\\");
        //set log directory, optional 
        ManualLog.setBasicPathAndInit("D:\\tmpdel\\gridwebjavatmp\\");
        SpringApplication.run(Application.class, args);
        System.out.print("helllo jdk ");
    }
    
 
}