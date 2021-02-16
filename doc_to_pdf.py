import os
import comtypes.client

wdFormatPDF = 17  # number of format in M.Word application
counter = 0

input_dir = r'C:\Users\alekz\OneDrive\Рабочий стол\doc_to_pdf\Архив ПТМ'
output_dir = r'C:\Users\alekz\OneDrive\Рабочий стол\doc_to_pdf\Архив ИЦ'

for dirs, dirs_name, files in os.walk(input_dir):
    try:
        os.stat(dirs.replace('Архив ПТМ', 'Архив ИЦ'))  # check specified directory
    except:
        os.mkdir(dirs.replace('Архив ПТМ', 'Архив ИЦ'))  # create a directory if dosnt have
    for file in files:
        check_file = file.split('.')[1]
        if check_file == 'doc' or check_file == 'docx':
            print('Найден файл:', file)
            print('Идет процесс конвертации в PDF')
            output_file = file.split('.')[0]  # splits the str by 2, makes a list, and passes the first str to the variable
            out_file = dirs + '\\' + output_file + '.pdf'
            word = comtypes.client.CreateObject('Word.Application', dynamic=True)  # without "dynamic" dont work
            word.Visible = False  # works together with both True and False
            in_file = os.path.join(dirs, file)
            doc = word.Documents.Open(in_file)
            doc.SaveAs(out_file.replace('Архив ПТМ', 'Архив ИЦ'), wdFormatPDF)
            doc.Close()
            word.Quit()
            print('Создан файл:', out_file.replace('Архив ПТМ', 'Архив ИЦ'))
            counter += 1
            print('______________')
print('Success! Program converted', counter, 'files')
print(input('Press ENTER to exit'))


            # todo:
            #0. Сделать GUI
            #1. Сделать прогресс бар
            #2. Прикрутить кнопку Отмена
            #3. Сделать exe