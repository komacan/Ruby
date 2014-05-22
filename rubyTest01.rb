require "spreadsheet"

book = Spreadsheet.open 'mysheet2.xls'
sheet1 = book.worksheet 0
rowNum = sheet1.row_count

puts rowNum
#puts sheet1.rows
#puts sheet1.cell(0,0)

rowNum = sheet1.row_count
colNum = sheet1.column_count

for i in 0..rowNum-1
	for j in 0..colNum-1
		puts sheet1.cell(i,j)
	end
end

