require 'roo'

def modifyUsersFile(filename)
    fio = filename.sheet(0).column(1).drop(1)
    logins = filename.sheet(0).column(2).drop(1)

    i = 0
    while i < fio.length do
        File.write('./users.txt', "#{fio[i]}\t#{logins[i]}\n", mode: 'a')
        i += 1
    end
    puts "Ваш файл users.txt готов! Для выхода нажмите ENTER.\nYour users.txt file is ready! Press ENTER to exit."
    exit = gets
end

xlsx = Roo::Spreadsheet.open('./docs/users_template.xlsx')
xlsx = Roo::Excelx.new("./docs/users_template.xlsx")
if File.exist?('users.txt')
    puts "Файл users.txt уже существует и будет дополнен. Продолжить? (y/n)"
    decision = gets.chomp
    if decision == "y"
        modifyUsersFile(xlsx)
    else
        puts "Пожалуйста удалите текущий файл users.txt из директории и перезапустите скрипт. Нажмите ENTER для выхода.\nPlease delete current users.txt file from directory and restart script. Press ENTER to exit."
        exit = gets
    end
else
    puts "Создается новый файл пользователей...\nCreating new users file..."
    modifyUsersFile(xlsx)
end
