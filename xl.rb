require "selenium-webdriver"
require "spreadsheet"

book = Spreadsheet.open 'mysheet.xls'
sheet1 = book.worksheet 0

puts sheet1.row_count

puts sheet1.last_row

puts sheet1.last_row_index

puts "After this line loop"

sheet1.each do |rows|
    puts rows[0]
end


#row = sheet1.row(0)
#puts row[0]





i=0	
sheet1.each do |rows|
		line = sheet1.row(rows)
		puts line
		#puts rows[0]
		#puts rows[1]
		#sheet1.rows[2].push "a"
	
	
	#increment in row
		#i += 1
end
#book.write 'mysheet3.xls'


