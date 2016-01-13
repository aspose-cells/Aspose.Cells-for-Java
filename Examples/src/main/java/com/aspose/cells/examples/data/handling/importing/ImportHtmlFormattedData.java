package com.aspose.cells.examples.data.handling.importing;

import com.aspose.cells.Cells;
import com.aspose.cells.ImportTableOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

import java.util.ArrayList;

public class ImportHtmlFormattedData {

    public static void main(String[] args) throws Exception {
        String dataDir = Utils.getDataDir(ImportHtmlFormattedData.class);
        String output1File = "Output.xlsx";
        String output2File = "Output.ods";
        String output1Path = dataDir + output1File;
        String output2Path = dataDir + output2File;

        ImportTableOptions importTableOptions = new ImportTableOptions();
        importTableOptions.setHtmlString(true);

        ArrayList<Employee> list = new ArrayList<>();
        list.add(new Employee()
                .setFirstName("a")
                .setLastName("a")
                .setAddress("a")
                .setCity("a")
                .setCountry("a")
        );
        list.add(new Employee()
                .setFirstName("<b>b</b>")
                .setLastName("b")
                .setAddress("b")
                .setCity("b")
                .setCountry("b")
        );

        Workbook workbook = new Workbook();
        Cells cells = workbook.getWorksheets().get(0).getCells();
        cells.importCustomObjects(list, 0, 0, importTableOptions);

        workbook.save(output1Path);
        System.out.println("File saved " + output1Path);
        workbook.save(output2Path);
        System.out.println("File saved " + output2Path);
    }
}

class Employee {
    private String firstName;
    private String lastName;
    private String address;
    private String city;
    private String country;

    public String getFirstName() {
        return firstName;
    }

    public Employee setFirstName(String firstName) {
        this.firstName = firstName;
        return this;
    }

    public String getLastName() {
        return lastName;
    }

    public Employee setLastName(String lastName) {
        this.lastName = lastName;
        return this;
    }

    public String getAddress() {
        return address;
    }

    public Employee setAddress(String address) {
        this.address = address;
        return this;
    }

    public String getCity() {
        return city;
    }

    public Employee setCity(String city) {
        this.city = city;
        return this;
    }

    public String getCountry() {
        return country;
    }

    public Employee setCountry(String country) {
        this.country = country;
        return this;
    }
}
