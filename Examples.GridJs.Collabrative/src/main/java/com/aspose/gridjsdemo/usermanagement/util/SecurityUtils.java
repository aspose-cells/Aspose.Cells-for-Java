package com.aspose.gridjsdemo.usermanagement.util;

import java.util.jar.JarFile;

import org.springframework.core.io.Resource;
import org.springframework.core.io.support.PathMatchingResourcePatternResolver;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.stereotype.Component;

import com.aspose.gridjsdemo.usermanagement.dto.CustomUserDetails;

@Component
public class SecurityUtils {
    
    public static CustomUserDetails getCurrentUser() {
        Authentication authentication = SecurityContextHolder.getContext().getAuthentication();
        if (authentication != null && authentication.getPrincipal() instanceof CustomUserDetails) {
            return (CustomUserDetails) authentication.getPrincipal();
        }
        throw new IllegalStateException("User not authenticated");
    }
    
	private static void printClassInJars() {
		System.out.println(getJarPath(org.springframework.data.annotation.CreatedDate.class));
    	System.out.println(getJarPath(javax.persistence.Table.class));
    	System.out.println(getJarPath(org.springframework.data.jpa.domain.support.AuditingEntityListener.class));
    	System.out.println(getJarPath(com.fasterxml.jackson.annotation.JsonFormat.class));
    	System.out.println(getJarPath(org.springframework.messaging.simp.SimpMessagingTemplate.class));
    	System.out.println(getJarPath(java.time.LocalDateTime.class));
    	System.out.println(getJarPath(org.springframework.beans.factory.aspectj.ConfigurableObject.class));
    	System.out.println(getJarPath(com.aspose.cells.Cells.class));
    	System.out.println(getJarPath(com.aspose.gridjs.OprMessageService.class));
    	System.out.println(getJarPath(org.springframework.boot.autoconfigure.condition.ConditionalOnProperty.class));
	}
	
    public static String getJarPath(Class<?> clazz) {
    	String ret=null;
        ClassLoader classLoader = clazz.getClassLoader();
        if (classLoader == null) {
            classLoader = ClassLoader.getSystemClassLoader();
        }

        String classResource = clazz.getName().replace(".", "/") + ".class";
        java.net.URL resource = classLoader.getResource(classResource);
        if (resource == null) return null;

        String url = resource.toString();
        if (url.startsWith("jar:file:")) {
            int exclamation = url.indexOf("!");
            ret= url.substring(4, exclamation); // 去掉 jar: 并截取到 ! 之前
        } else if (url.startsWith("file:")) {
            // 类可能在开发目录中（未打包）
            ret= "Class is in classpath (not in JAR): " + url;
        }
        ret= url;
        return ret.replace('/', '\\');
    }
    
	private static void debugAnotation() {
		try {
            Class<?> clazz = Class.forName("com.aspose.gridjs.OprMessageService");
            boolean hasService = clazz.isAnnotationPresent(org.springframework.stereotype.Service.class);
            System.out.println("Has @Service: " + hasService);

            // 打印所有注解
            System.out.println("All annotations:");
            for (java.lang.annotation.Annotation ann : clazz.getAnnotations()) {
                System.out.println("  " + ann);
            }
        } catch (ClassNotFoundException e) {
            System.out.println("Class not found!");
        }
	}
	
	private static void debugSpring() {
		try {
    	    Class<?> clazz = Class.forName("com.aspose.gridjs.OprMessageRepository");
    	    System.out.println("✅ 类存在: " + clazz);
    	    System.out.println("✅ 有 @Repository: " + clazz.isAnnotationPresent(org.springframework.stereotype.Repository.class));
    	    
    	    
	    	    PathMatchingResourcePatternResolver resolver = new PathMatchingResourcePatternResolver();
	            Resource[] resources = resolver.getResources("classpath*:com/aspose/gridjs/**/*.class");
	            System.out.println("Found " + resources.length + " resources");
	            for (Resource res : resources) {
	                System.out.println(res.getURL());
	            }
	            
 
	            
	            try (JarFile jar = new JarFile("D:\\release\\v25.8\\aspose-cells-25.8-java\\aspose-gridjs-25.8-fix.jar")) {
	                jar.stream()
	                    .filter(e -> e.isDirectory())
	                    .forEach(e -> System.out.println("DIR: " + e.getName()));
	            }
	            
	            
	            // ConfigurableApplicationContext context = SpringApplication.run(UserManagementApplication.class, args);
	         
	             /*
	              // 打印所有 Bean 名称，查找 OprMessageService
	              String[] beanNames = context.getBeanDefinitionNames();
	              Arrays.stream(beanNames)
	                     .filter(name -> (name.contains("OprMessage")||name.contains("GridJs")))
	                   .forEach(name -> System.out.println("🔍 Found Bean: " + name + " -> " + context.getBean(name).getClass()));
	         
	            */
	            	
	            	
//	         	    printClassInJars();
	            	
//	            	debugAnotation();
	            	 
 
 

            
            
    	} catch (Exception e) {
    	    System.out.println("❌ 类未找到！");
    	}
	}
}

// 使用方式：
//CustomUserDetails user = SecurityUtils.getCurrentUser();
