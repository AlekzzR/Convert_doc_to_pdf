import os
import tkinter as tk
import comtypes.client
import tkinter.filedialog as fd

wdFormatPDF = 17  # number of format in M.Word application
counter = 0

class MyGUI:
    def __init__(self):
        self.main_window = tk.Tk()
        self.label1 = tk.Label(self.main_window, text='  ')
        self.label2 = tk.Label(self.main_window, text='Программа для конвертации файлов doc в pdf')
        self.label3 = tk.Label(self.main_window, text='  ')
        self.label4 = tk.Label(self.main_window, text='  ')
        self.button_dir = tk.Button(self.main_window, text='Выбрать папку', command=self.choose_directory)
        self.button_start = tk.Button(self.main_window, text='Погнали', command=self.loop)
        self.label2.pack()
        self.label1.pack()
        self.button_dir.pack()
        self.label3.pack()
        self.button_start.pack()
        self.label4.pack()
        tk.mainloop()

    def choose_directory(self):
        self.directory = fd.askdirectory(title="Открыть папку", initialdir="/")
        if self.directory:
            print(self.directory)
            return self.directory

    def loop(self):
        loop_pdf(self.directory)

def loop_pdf(dir):
   for dirs, dirs_name, files in os.walk(dir.replace('/', '\\')):
    try:
        os.stat(dirs.replace('Архив ПТМ', 'Архив ИЦ'))  # check specified directory
    except:
        os.mkdir(dirs.replace('Архив ПТМ', 'Архив ИЦ'))  # create a directory if doesn't have
    for file in files:
        check_file = file.split('.')[1]
        if check_file == 'doc' or check_file == 'docx':
            print('Найден файл:', file)
            print('Идет процесс конвертации в PDF')
            output_file = file.split('.')[0]
            # splits the str by 2, makes a list, and passes the first str to the var
            out_file = dirs + '\\' + output_file + '.pdf'
            word = comtypes.client.CreateObject('Word.Application', dynamic=True)
            # without "dynamic" don't work
            word.Visible = False  # works together with both True and False
            in_file = os.path.join(dirs, file)
            doc = word.Documents.Open(in_file)
            doc.SaveAs(out_file.replace('Архив ПТМ', 'Архив ИЦ'), wdFormatPDF)
            doc.Close()
            word.Quit()
            print('Создан файл:', out_file.replace('Архив ПТМ', 'Архив ИЦ'))
            # counter += 1
            print('______________')
#input_dir = r'C:\Users\alekz\OneDrive\Рабочий стол\doc_to_pdf\Архив ПТМ'
#output_dir = r'C:\Users\alekz\OneDrive\Рабочий стол\doc_to_pdf\Архив ИЦ'



my_gui = MyGUI()
#print('Success! Program converted', counter, 'files')
#print(input('Press ENTER to exit'))

# todo:
# 0. Вывести информационное диалоговое окно по окончании со статистикой
# 1. Сделать прогресс бар
# 2. Прикрутить кнопку Отмена
