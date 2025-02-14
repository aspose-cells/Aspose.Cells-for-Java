# Aspose.Cells for Java
[Aspose.Cells.GridJs for Java](http://www.aspose.com/products/cells/java)  is a lightweight, scalable, and customizable toolkit that provides cross-platform web applications, enables convenient development for editing or viewing Excel/Spreadsheet files, offers simple deployment, and provides easy-to-use APIs.

## Preview

 <img alt='demo' src='https://unpkg.com/gridjs-spreadsheet@25.1.0/preview.gif' />

## How to Run the Examples
###1.edit the properties in src\main\resources\application.properties to meet your local environment
```properties
# This is the directory contains the spread sheet files
testconfig.ListDir=/app/wb

# Directory for storing cache files
testconfig.CachePath=/app/grid_cache

# Directory for storing upload files
testconfig.UploadPath=/app/upload

# License file for Aspose.Cells
testconfig.AsposeLicensePath=/app/license
```
###2.run src\main\java\com\aspose\gridjs\demo\GridjsdemoApplication.java

open browser and navigate to view all the files in the directory at http://localhost:8080/gridjsdemo/list

## Step to run in docker 

1. docker build -t gridjs-demo-java .

2. run with aspose license file:
      docker run -d -p 8080:8080  -v C:/path/to/license.txt:/app/license gridjs-demo-java
   or just run the demo in trial mode:
      docker run -d -p 8080:8080 gridjs-demo-java
      
3. open browser and enter the url: http://localhost:8080/gridjsdemo/list

## Resources

+ **Website:** [www.aspose.com](http://www.aspose.com) 
+ **Product Home:** [Aspose.Cells for Java](http://www.aspose.com/products/cells/java)
+ **Download:** [Download Aspose.Cells for Java](https://downloads.aspose.com/cells/java)
+ **Documentation:** [Aspose.Cells for Java Documentation](https://docs.aspose.com/display/cellsjava/Home)
+ **Forum:** [Aspose.Cells for Java Forum](http://www.aspose.com/community/forums/aspose.cells-product-family/19/showforum.aspx)
+ **Blog:** [Aspose.Cells for Java Blog](https://blog.aspose.com/category/aspose-products/aspose-cells-product-family/)
