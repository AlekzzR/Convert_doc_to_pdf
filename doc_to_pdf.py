import os
import tkinter as tk
import tkinter.messagebox
import comtypes.client
import tkinter.filedialog as fd


wdFormatPDF = 17  # number of format in M.Word application

class MyGUI:
    def __init__(self):
        self.counter = 0
        self.main_window = tk.Tk()
        self.label1 = tk.Label(self.main_window, text='  ')
        self.label2 = tk.Label(self.main_window, text='Программа для конвертации файлов doc в pdf')
        self.label3 = tk.Label(self.main_window, text='  ')
        self.label4 = tk.Label(self.main_window, text='  ')
        self.label5 = tk.Label(self.main_window, text='  ')
        self.value = tk.StringVar()
        self.label6 = tk.Label(self.main_window, textvariable=self.value)
        self.button_dir = tk.Button(self.main_window, text='  Выбрать папку  ', command=self.choose_directory, background="#999")
        self.button_start = tk.Button(self.main_window, text='  Конвертировать в PDF  ', command=self.loop, background="#789", font="12")
        self.button_exit = tk.Button(self.main_window,text='  Выход  ', command=self.main_window.destroy, background="#988")
        self.label2.pack()
        self.label1.pack()
        self.button_dir.pack()
        self.label6.pack()
        self.label3.pack()
        self.button_start.pack()
        self.label4.pack()
        self.button_exit.pack()
        self.label5.pack()
        tk.mainloop()

    def choose_directory(self):
        self.directory = fd.askdirectory(title="Открыть папку", initialdir="/")
        self.value.set(self.directory)
        if self.directory:
            print(self.directory)
            return self.directory

    def loop(self):
       try:
           for dirs, dirs_name, files in os.walk(self.directory.replace('/', '\\')):
               try:
                   os.stat(dirs.replace('AРХИВ ООО ПТМ', 'Архив ИЦ'))  # check specified directory
               except:
                   os.mkdir(dirs.replace('AРХИВ ООО ПТМ', 'Архив ИЦ'))  # create a directory if doesn't have
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
                       doc.SaveAs(out_file.replace('AРХИВ ООО ПТМ', 'Архив ИЦ'), wdFormatPDF)
                       doc.Close()
                       word.Quit()
                       print('Создан файл:', out_file.replace('AРХИВ ООО ПТМ', 'Архив ИЦ'))
                       self.counter += 1
                       print('______________')
           self.show_info()
       except AttributeError:
           self.show_error()
    def show_info(self):
       tk.messagebox.showinfo('Информация', 'Преобразовано файлов: ' + str(self.counter))
    def show_error(self):
       tk.messagebox.showinfo('Ошибка', 'Не выбрана папка с файлами')



my_gui = MyGUI()
#print('Success! Program converted', counter, 'files')
#print(input('Press ENTER to exit'))

# todo:
# 1. Сделать прогресс бар
# 2. Сделать StringVar на конвертируемые файлы JIT
