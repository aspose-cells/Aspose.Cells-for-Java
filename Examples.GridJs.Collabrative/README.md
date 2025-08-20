[Product Page](https://products.aspose.com/cells/java) | [Docs](https://docs.aspose.com/cells/java/aspose-cells-gridjs/) | [API Reference](https://reference.aspose.com/cells/java/com.aspose.gridjs/) | [Demos](https://products.aspose.app/cells/family/) | [Blog](https://blog.aspose.com/category/cells/) | [Code Samples](https://github.com/aspose-cells/Aspose.Cells-for-Java/tree/master/Examples.GridJs) | [Free Support](https://forum.aspose.com/c/cells) | [Temporary License](https://purchase.aspose.com/temporary-license) | [EULA](https://company.aspose.com/legal/eula)  

---

Try our [free online apps](https://products.aspose.app/cells/family) demonstrating some of the most popular Aspose.Cells functionality.

[Aspose.Cells.GridJs for Java](http://www.aspose.com/products/cells/java) is a lightweight, scalable, and customizable toolkit that provides cross-platform web applications. It enables developers to easily build Excel/Spreadsheet editors or viewers for the web with collaborative features, simple deployment, and easy-to-use APIs.

This repository contains a demo project showing how to implement a **Spreadsheet Editor in collaborative mode** using **GridJs**.  

---

## Preview

<img alt="demo" src="https://unpkg.com/gridjs-spreadsheet@25.1.0/preview.gif" />

---

## Quick Start

1. **Clone the Repository**
   ```bash
   git clone https://github.com/aspose-cells/Aspose.Cells-for-Java.git
   cd Aspose.Cells-for-Java/Examples.GridJs.Collabrative
   ```

2. **Configure Application**
   - Update `src/main/resources/application.properties` with your environment.  
   - Ensure MySQL is installed and running.  

3. **Run the Demo**
   ```bash
   mvn spring-boot:run
   ```
   Open [http://localhost:8080/gridjsdemo/list](http://localhost:8080/gridjsdemo/list)

---

## Run the Example Locally

### 0. Install SQL Server
You will need a database server (e.g., **MySQL**). Make sure it is running before you start the application.

### 1. Configure Application Properties
Edit `src/main/resources/application.properties` according to your environment:

```properties
# Directory containing spreadsheet files
testconfig.ListDir=/app/wb

# Directory for storing cache files
testconfig.CachePath=/app/grid_cache

# Aspose.Cells license file
testconfig.AsposeLicensePath=/app/license

# Enable collaborative mode
gridjs.iscollabrative=true

# Database connection (example: MySQL)
spring.datasource.url=jdbc:mysql://localhost:3306/gridjsdemodb?createDatabaseIfNotExist=true&useUnicode=true&useJDBCCompliantTimezoneShift=true&useLegacyDatetimeCode=false&serverTimezone=Asia/Jakarta&useSSL=false
spring.datasource.username=root
spring.datasource.password=root
spring.sql.init.platform=mysql
```

### 2. Run the Application
Execute:

```
src/main/java/com/aspose/gridjsdemo/usermanagement/UserManagementApplication.java
```

Then open your browser and navigate to:  
  [http://localhost:8080/gridjsdemo/list](http://localhost:8080/gridjsdemo/list)

---

## Run with Docker

### 0. Configure License Path
In `docker-compose.yml` (line 10), set the correct path to your license file.

Example: If your license file is at `C:/license/aspose.lic`

Before:
```yaml
- D:/release/license/Aspose.Cells.lic:/app/license  # optional: set Aspose license file
```

After:
```yaml
- C:/license/aspose.lic:/app/license  # optional: set Aspose license file
```

### 1. Build and Start Containers
```bash
docker-compose up --build
```

### 2. Access the App
Open your browser and go to:  
  [http://localhost:8080/gridjsdemo/list](http://localhost:8080/gridjsdemo/list)

---

## JavaScript/CSS Libraries Used
- jquery.js v2.1.1  
- jquery-ui.js v1.12.1  
- jquery-ui.css v1.12.1  
- jszip.min.js v3.6.0  
- bootstrap.css v22.5.5.2  
- quantumui.css v22.5.5.2  

---

## Resources
- **Website:** [www.aspose.com](http://www.aspose.com)  
- **Product Home:** [Aspose.Cells for Java](http://www.aspose.com/products/cells/java)  
- **Download:** [Download Aspose.Cells for Java](https://downloads.aspose.com/cells/java)  
- **Documentation:** [Aspose.Cells for Java Documentation](https://docs.aspose.com/display/cellsjava/Home)  
- **Forum:** [Aspose.Cells for Java Forum](http://www.aspose.com/community/forums/aspose.cells-product-family/19/showforum.aspx)  
- **Blog:** [Aspose.Cells for Java Blog](https://blog.aspose.com/category/aspose-products/aspose-cells-product-family/)  
