require 'nokogiri'
load 'property.rb'

@properties = Array.new
Property::PROPERTY_NAMES.each {|key, value|
  property_variable = key.downcase.gsub(/ /, '')
  property = instance_variable_set("@#{property_variable}" , Property.new(key, value))
  @properties << property
}

# The logic in this portion of the code relates to data contained in the 50F Summary.
@doc = Nokogiri.XML(File.read("/Users/lcurley/Dropbox/dashboard_work/weekly_report_date/#{Time.now.strftime('%Y_%m_%d')}/50f_weekly_report/50f_summary.xml"))
(9..47).each {|row_number|
  property_code = @doc.xpath("/ss:Workbook/ss:Worksheet[1]/ss:Table[1]/ss:Row[#{row_number}]/ss:Cell[1]/ss:Data[1]/html:Font[1]/text()").to_s
  if property_code != "bs"
	  property = @properties.find_all {|p| p.alternate_names.include?("#{property_code}")}.first
	  if property
	  	property.phases += 1 
	  	property.total_units += @doc.xpath("/ss:Workbook/ss:Worksheet[1]/ss:Table[1]/ss:Row[#{row_number}]/ss:Cell[3]/ss:Data[1]/text()").to_s.to_i 
	  	property.current_occupied += @doc.xpath("/ss:Workbook/ss:Worksheet[1]/ss:Table[1]/ss:Row[#{row_number}]/ss:Cell[7]/ss:Data[1]/text()").to_s.to_i
	  	property.total_vacants += @doc.xpath("/ss:Workbook/ss:Worksheet[1]/ss:Table[1]/ss:Row[#{row_number}]/ss:Cell[8]/ss:Data[1]/text()").to_s.to_i
	  	property.vacant_rented += @doc.xpath("/ss:Workbook/ss:Worksheet[1]/ss:Table[1]/ss:Row[#{row_number}]/ss:Cell[10]/ss:Data[1]/text()").to_s.to_i
	  	property.vacant_unrented += @doc.xpath("/ss:Workbook/ss:Worksheet[1]/ss:Table[1]/ss:Row[#{row_number}]/ss:Cell[11]/ss:Data[1]/text()").to_s.to_i
	  	property.percent_preleased += @doc.xpath("/ss:Workbook/ss:Worksheet[1]/ss:Table[1]/ss:Row[#{row_number}]/ss:Cell[13]/ss:Data[1]/html:Font[1]/text()").to_s.gsub(/%/, '').to_f
  	end
  end
}

@properties.each {|property|
	property.percent_occupied = property.current_occupied.to_f / property.total_units.to_f
	property.percent_preleased = (property.percent_preleased.to_f / property.phases) / 100
}

# The logic in this portion of the code relates to data contained in the 81F report.
files = Dir.entries("/Users/lcurley/Dropbox/dashboard_work/weekly_report_date/#{Time.now.strftime('%Y_%m_%d')}/81f_tour_app").delete_if {|x| x == '.' or x == '..' or /pdf/ =~ x}
files.each {|file|
	@doc = Nokogiri.XML(File.read("/Users/lcurley/Dropbox/dashboard_work/weekly_report_date/#{Time.now.strftime('%Y_%m_%d')}/81f_tour_app/#{file}"))
	property_code = @doc.xpath("/ss:Workbook/ss:Worksheet[1]/ss:Table[1]/ss:Row[4]/ss:Cell[1]/ss:Data[1]/html:Font[1]/text()").to_s
	property_code = property_code.gsub(/\n/, '').squeeze(' ')
	property = @properties.find_all {|p| p.alternate_names.include?("#{property_code}")}.first
	row_number = @doc.xpath("count(/ss:Workbook/ss:Worksheet[1]/ss:Table[1]/ss:Row)").to_i
	property.total_guest_cards += @doc.xpath("/ss:Workbook/ss:Worksheet[1]/ss:Table[1]/ss:Row[#{row_number - 1}]/ss:Cell[3]/ss:Data[1]/text()").to_s.to_i
	property.total_apps += @doc.xpath("/ss:Workbook/ss:Worksheet[1]/ss:Table[1]/ss:Row[#{row_number - 1}]/ss:Cell[4]/ss:Data[1]/text()").to_s.to_i
}