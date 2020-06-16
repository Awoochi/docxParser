require 'docx'

def buildTable(table)

  skus = Array.new
  pallets = Array.new
  mdates = Array.new
  lots = Array.new
  qtys = Array.new
  newDates = Array.new

  table.columns.each_with_index do |col,ind|
    if ind != 5
      col.cells.each do |cell|
        if !cell.text.empty? && /\d+/.match(cell.text)
          case ind
            when 0
              skus.push(cell.text)
            when 1
              pallets.push(cell.text)
            when 2
              mdates.push(cell.text)
            when 3
              lots.push(cell.text)
            when 4
              qtys.push(cell.text)
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

doc = Docx::Document.open('./docs/receives.docx')

table1 = doc.tables[0]
table2 = doc.tables[1]
buildTable(table1)
buildTable(table2)