require "selenium-webdriver"
require "spreadsheet"

book = Spreadsheet.open 'mysheet.xls'
sheet1 = book.worksheet 0
rowCount = sheet1.row_count
puts "Total iteration : #{rowCount}"
driver = Selenium::WebDriver.for :firefox

#Initialize variables for loop
rowNum = sheet1.row_count
colNum = sheet1.column_count
i=0	
for i in 0..rowNum-1
		data1 = sheet1.cell(i,0)
	#Write log Iteration Number
		puts "Iteration Number : #{i+1}"
		puts "Iteration IP : #{data1}"
	#Firefox browser instantiation
		driver.navigate.to "http://www.yahoo.com"
	#SearchBox.send_keys data1
		driver.find_element(:id, "p_13838465-p").send_keys data1
	#Clicking Search Button
		driver.find_element(:id, "search-submit").click
	#Set the row number in sheet1
		row = sheet1.row(i)
	#Write to intersection of (i)th row and 2nd column in sheet1
		row[2] = driver.title
	#Write log whether title matches the expected value
		puts (if "#{data1} - Yahoo Search Results" == driver.title
			"Search completed with keyword"
		else
			"Error Found!"
		end)
		puts "Page title : #{driver.title}",""
end
book.write 'mysheet2.xls'
#Quitting the browser
driver.quit

