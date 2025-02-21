[Product Page](https://products.aspose.com/cells/java) | [Docs](https://docs.aspose.com/cells/java/aspose-cells-gridjs/) | [API Reference](https://reference.aspose.com/cells/java/com.aspose.gridjs/) | [Demos](https://products.aspose.app/cells/family/) | [Blog](https://blog.aspose.com/category/cells/) | [Code Samples](https://github.com/aspose-cells/Aspose.Cells-for-Java/tree/master/Examples.GridJs) | [Free Support](https://forum.aspose.com/c/cells) | [Temporary License](https://purchase.aspose.com/temporary-license) | [EULA](https://company.aspose.com/legal/eula) 

Try our [free online apps](https://products.aspose.app/cells/family) demonstrating some of the most popular Aspose.Cells functionality.

[Aspose.Cells.GridJs for Java](http://www.aspose.com/products/cells/java)  is a lightweight, scalable, and customizable toolkit that provides cross-platform web applications, enables convenient development for editing or viewing Excel/Spreadsheet files, offers simple deployment, and provides easy-to-use APIs.

This is a  demo to show how we can use GridJs to implment a spreadsheet Editor .
 

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

```bash
      docker run -d -p 8080:8080  -v C:/path/to/license.txt:/app/license gridjs-demo-java
```

   or just run the demo in trial mode:
   
```bash
      docker run -d -p 8080:8080 gridjs-demo-java
```
      
3. open browser and enter the url: http://localhost:8080/gridjsdemo/list

## Reference js lib used in the demo project:
jquery.js    v2.1.1

jquery-ui.js v1.12.1 

jquery-ui.css v1.12.1 

jszip.min.js v3.6.0 

bootstrap.css   v22.5.5.2

quantumui.css   v22.5.5.2

## Resources

+ **Website:** [www.aspose.com](http://www.aspose.com) 
+ **Product Home:** [Aspose.Cells for Java](http://www.aspose.com/products/cells/java)
+ **Download:** [Download Aspose.Cells for Java](https://downloads.aspose.com/cells/java)
+ **Documentation:** [Aspose.Cells for Java Documentation](https://docs.aspose.com/display/cellsjava/Home)
+ **Forum:** [Aspose.Cells for Java Forum](http://www.aspose.com/community/forums/aspose.cells-product-family/19/showforum.aspx)
+ **Blog:** [Aspose.Cells for Java Blog](https://blog.aspose.com/category/aspose-products/aspose-cells-product-family/)
