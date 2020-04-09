require 'docx'

#first_table = doc.tables[0]
#puts first_table.row_count
#puts first_table.column_count
#puts first_table.rows[0].cells[0].text
#puts first_table.columns[0].cells[0].text
=begin
# Iterate through tables
doc.tables.each do |table|
  table.rows.each do |row| # Row-based iteration
    row.cells.each do |cell|
      puts cell.text
    end
  end

  table.columns.each do |column| # Column-based iteration
    column.cells.each do |cell|
      puts cell.text
    end
  end
end
=end

#table1.columns[0].cells.each_with_index do |cell,ind|
#    if ind == 0 || ind == 1 || ind == 2
#        puts cell
#    end
#end

doc = Docx::Document.open('receives.docx')

def readTable(table) #deprecated method
sku = ""
pallet = ""
mdate = ""
lot = ""
qty = ""
  table.rows.each_with_index do |row,index|
        if index >= 2 && table.columns.index != 5
          p index
              table.rows[1].cells.each_with_index do |c,i|
                    case i
                        when 0
                            sku = c.text
                            p sku
                        when 1
                            pallet = c.text
                            p pallet
                        when 2
                            mdate = c.text
                            p mdate
                        when 3
                            lot = c.text
                            p lot
                        when 4
                            qty = c.text
                            p qty
                        else
                            puts ""
                    end
              end
          mdate = dateFormat(mdate)
          puts "240#{sku}\\x1d243#{pallet}\\x1d11#{mdate}10#{lot}\\x1d37#{qty}" 
          puts "This is date: #{mdate}" #02.04.2020
        end
  end
end

def buildTable(table)
skus = Array.new
pallets = Array.new
mdates = Array.new
lots = Array.new
qtys = Array.new
newDates = Array.new

  table.columns.each_with_index do |col,ind|
#    p "=========================COLUMN #{ind}"
    if ind != 5
      col.cells.each do |cell|
        if !cell.text.empty? && /\d+/.match(cell.text)
#          p cell.text
          case ind
            when 0
              skus.push(cell.text)
#              p skus
            when 1
              pallets.push(cell.text)
#              p pallets
            when 2
              mdates.push(cell.text)
#              p mdates
            when 3
              lots.push(cell.text)
#              p lots
            when 4
              qtys.push(cell.text)
#              p qtys
            else
              puts ""
          end
        end
      end
    end
  end
  mdates.each do |date|
    newDate = dateFormat(date)
    newDates.push(newDate)
  end
createCodes(skus,pallets,newDates,lots,qtys)
end

def dateFormat(date)
  day = date[0..1]
  month = date[3..4]
  year = date[8..9]

  return year + month + day
end

def createCodes(skus,pallets,dates,lots,qtys)
  i = 0
  while i < skus.length do
    File.write('codes.txt',"240#{skus[i]}\\x1d243#{pallets[i]}\\x1d11#{dates[i]}10#{lots[i]}\\x1d37#{qtys[i]}\n", mode: 'a')  
    i += 1
  end
end

table1 = doc.tables[0]
table2 = doc.tables[1]
buildTable(table1)
buildTable(table2)











#p table1.row_count
#p table1.rows[3].cells[0].text

=begin
  table1.rows.each_with_index do |row, ind| # Row-based iteration
    p "=======================================ROW NUMBER #{ind}"
    row.cells.each do |cell|
      puts cell.text
    end
  end

  table1.columns.each_with_index do |column, index| # Column-based iteration
    p "============================================COLUMN #{index}"
    column.cells.each do |cell|
      puts cell.text
    end
  end
=end
#readTable(table1)