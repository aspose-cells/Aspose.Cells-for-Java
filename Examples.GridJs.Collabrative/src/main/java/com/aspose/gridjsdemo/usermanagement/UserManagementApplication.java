package com.aspose.gridjsdemo.usermanagement;

import java.util.jar.JarFile;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.autoconfigure.domain.EntityScan;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.ComponentScan;
import org.springframework.core.io.Resource;
import org.springframework.core.io.support.PathMatchingResourcePatternResolver;
import org.springframework.data.jpa.repository.config.EnableJpaAuditing;
import org.springframework.data.jpa.repository.config.EnableJpaRepositories;
import org.springframework.security.crypto.bcrypt.BCryptPasswordEncoder;
import org.springframework.security.crypto.password.PasswordEncoder;

@SpringBootApplication
@EnableJpaAuditing
@ComponentScan(basePackages = { "com.aspose.gridjs","com.aspose.gridjsdemo"})
@EnableJpaRepositories(basePackages = {"com.aspose.gridjs","com.aspose.gridjsdemo"})  
@EntityScan(basePackages = {"com.aspose.gridjs","com.aspose.gridjsdemo.usermanagement.entity"})  
public class UserManagementApplication {

    public static void main(String[] args) {
    	
    	 
    	
    	
    	
    SpringApplication.run(UserManagementApplication.class, args);
     
    }
    
    
    
  
}
