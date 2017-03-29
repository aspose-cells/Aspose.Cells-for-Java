# Aspose_Cells_Java_for_Ruby
Aspose Cells Java for Ruby is a gem that demonstrates / provides the Aspose.Cells for Java API usage examples in Ruby by using Rjb - Ruby Java Bridge.

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'asposecellsjava'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install asposecellsjava

To download Aspose.Cells for Java API to be used with these examples through RJB, Please navigate to:

https://downloads.aspose.com/cells/java

For most complete documentation of the project, check Aspose.Cells Java for Ruby confluence wiki link:

https://docs.aspose.com/display/cellsjava/Aspose.Cells+Java+For+Ruby

## Usage

```ruby
require require File.dirname(File.dirname(File.dirname(__FILE__))) + '/lib/asposecellsjava'
include Asposecellsjava
include Asposecellsjava::HelloWorld
initialize_aspose_cells
```
Lets understand the above code
* The first line makes sure that the aspose cells is loaded and available 
* Include the files that are required to access the aspose cells
* Initialize the libraries. The aspose JAVA classes are loaded from the path provided in the aspose.yml file
