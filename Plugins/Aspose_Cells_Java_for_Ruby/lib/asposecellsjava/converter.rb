module Asposecellsjava
  module Converter
    def initialize()
        @data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(@data_dir + 'Book1.xls')
        
        # Converting Excel to PDF
        excel_to_pdf(workbook)

        # Converting Chart to Image
        chart_to_image()

        # Converting Worksheet to Image
        worksheet_to_image(workbook)

        # Converting Worksheet to SVG
        worksheet_to_svg(workbook)

        # Converting Worksheet to MHTML
        worksheet_to_mhtml(workbook)

        # Converting Worksheet to HTML
        worksheet_to_html(workbook)

        # Converting HTML to Excel
        html_to_excel()
    end

    def excel_to_pdf(workbook)
        save_format = Rjb::import('com.aspose.cells.SaveFormat')

        # Save the document in PDF format
        workbook.save(@data_dir + "MyPdfFile.pdf", save_format.PDF)

        puts "Pdf saved successfully."
    end    

    def chart_to_image()
        # Create a new Workbook.
        workbook = Rjb::import('com.aspose.cells.Workbook').new

        # Get the first worksheet.
        sheet = workbook.getWorksheets().get(0)

        # Set the name of worksheet
        sheet.setName("Data")

        # Get the cells collection in the sheet.
        cells = workbook.getWorksheets().get(0).getCells()

        # Put some values into a cells of the Data sheet.
        cells.get("A1").setValue("Region")
        cells.get("A2").setValue("France")
        cells.get("A3").setValue("Germany")
        cells.get("A4").setValue("England")
        cells.get("A5").setValue("Sweden")
        cells.get("A6").setValue("Italy")
        cells.get("A7").setValue("Spain")
        cells.get("A8").setValue("Portugal")
        cells.get("B1").setValue("Sale")
        cells.get("B2").setValue(70000)
        cells.get("B3").setValue(55000)
        cells.get("B4").setValue(30000)
        cells.get("B5").setValue(40000)
        cells.get("B6").setValue(35000)
        cells.get("B7").setValue(32000)
        cells.get("B8").setValue(10000)

        # Create chart
        chart_type = Rjb::import('com.aspose.cells.ChartType')
        chart_index = sheet.getCharts().add(chart_type.COLUMN, 12, 1, 33, 12)
        chart = sheet.getCharts().get(chart_index)

        # Set properties of chart title
        chart.getTitle().setText("Sales By Region")
        chart.getTitle().getFont().setBold(true)
        chart.getTitle().getFont().setSize(12)

        # Set properties of nseries
        chart.getNSeries().add("Data!B2:B8", true)
        chart.getNSeries().setCategoryData("Data!A2:A8")

        # Set the fill colors for the series's data points (France - Portugal(7 points))
        chart_points = chart.getNSeries().get(0).getPoints()

        color = Rjb::import('com.aspose.cells.Color')

        point = chart_points.get(0)
        point.getArea().setForegroundColor(color.getCyan())

        point = chart_points.get(1)
        point.getArea().setForegroundColor(color.getBlue())

        point = chart_points.get(2)
        point.getArea().setForegroundColor(color.getYellow())

        point = chart_points.get(3)
        point.getArea().setForegroundColor(color.getRed())

        point = chart_points.get(4)
        point.getArea().setForegroundColor(color.getBlack())

        point = chart_points.get(5)
        point.getArea().setForegroundColor(color.getGreen())

        point = chart_points.get(6)
        point.getArea().setForegroundColor(color.getMaroon())

        # Set the legend invisible
        chart.setShowLegend(false)

        # Get the Chart image
        img_opts = Rjb::import('com.aspose.cells.ImageOrPrintOptions').new
        image_format = Rjb::import('com.aspose.cells.ImageFormat')
        img_opts.setImageFormat(image_format.getPng())

        # Save the chart image file.
        chart.toImage(@data_dir + "MyChartImage.png", img_opts)

        # Print message
        puts "Convert chart to image successfully."
    end    

    def worksheet_to_image(workbook)
        #Create an object for ImageOptions
        img_options = Rjb::import('com.aspose.cells.ImageOrPrintOptions').new
        
        # Set the image type
        image_format = Rjb::import('com.aspose.cells.ImageFormat')
        img_options.setImageFormat(image_format.getPng())
        
        # Get the first worksheet.
        sheet = workbook.getWorksheets().get(0)

        # Create a SheetRender object for the target sheet
        sr = Rjb::import('com.aspose.cells.SheetRender').new(sheet, img_options)
        
        j = 0
        while j < sr.getPageCount()
            # Generate an image for the worksheet
            sr.toImage(j, @data_dir + "mysheetimg_#{j}.png")
            j +=1
        end

        puts "Image saved successfully."
    end  

    def worksheet_to_svg(workbook)
        # Convert each worksheet into svg format in a single page.
        img_options = Rjb::import('com.aspose.cells.ImageOrPrintOptions').new
        save_format = Rjb::import('com.aspose.cells.SaveFormat')
        img_options.setSaveFormat(save_format.SVG)
        img_options.setOnePagePerSheet(true)
        
        # Convert each worksheet into svg format
        sheet_count = workbook.getWorksheets().getCount()

        i=0
        while i < sheet_count
            sheet = workbook.getWorksheets().get(i)

            sr = Rjb::import('com.aspose.cells.SheetRender').new(sheet, img_options)

            k=0
            while sr.getPageCount()
                # Output the worksheet into Svg image format
                sr.toImage(k, @data_dir + sheet.getName() + "#{k}.svg")
            end
        end

        puts "SVG saved successfully."
    end  

    def worksheet_to_mhtml(workbook)
        save_format = Rjb::import('com.aspose.cells.SaveFormat')
        # Specify the HTML saving options
        sv = Rjb::import('com.aspose.cells.HtmlSaveOptions').new(save_format.M_HTML)

        # Save the document
        workbook.save(@data_dir + "convert.mht", sv)

        puts "MHTML saved successfully."
    end 

    def worksheet_to_html(workbook)
        save_format = Rjb::import('com.aspose.cells.SaveFormat')
        # Specify the HTML saving options
        save = Rjb::import('com.aspose.cells.HtmlSaveOptions').new(save_format.M_HTML)

        # Save the document
        workbook.save(@data_dir + "output.html", save)

        puts "HTML saved successfully."
    end  

    def html_to_excel()
        load_format = Rjb::import('com.aspose.cells.LoadFormat')
        # Create an instance of HTMLLoadOptions and initiate it with appropriate LoadFormat
        options = Rjb::import('com.aspose.cells.HTMLLoadOptions').new(load_format.HTML)
        
        # Load the Html file through file path while passing the instance of HTMLLoadOptions class
        workbook = Rjb::import('com.aspose.cells.Workbook').new(@data_dir + "index.html", options)
        
        save_format = Rjb::import('com.aspose.cells.SaveFormat')
        #Save the results to disc in Xlsx format
        workbook.save(@data_dir + "output.xlsx", save_format.XLSX)

        puts "XLSX saved successfully."
    end
  end
end
