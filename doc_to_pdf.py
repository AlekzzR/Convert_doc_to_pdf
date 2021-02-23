import os
import tkinter as tk
import tkinter.messagebox
import comtypes.client
import tkinter.filedialog as fd


wdFormatPDF = 17  # number of format in M.Word application

class MyGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.counter = 0
        self.label1 = tk.Label(self, text='  ')
        self.label2 = tk.Label(self, text='Программа для конвертации файлов doc в pdf')
        self.label3 = tk.Label(self, text='  ')
        self.label4 = tk.Label(self, text='  ')
        self.label5 = tk.Label(self, text='  ')
        self.value = tk.StringVar()
        self.label6 = tk.Label(self, textvariable=self.value)
        self.button_dir = tk.Button(self, text='  Выбрать папку  ', command=self.choose_directory,
                                    background="#f8f8ff", relief=tk.SOLID)
        self.button_start = tk.Button(self, text='  Конвертировать в PDF  ', command=self.loop,
                                      background="#f8f8ff", relief=tk.SOLID)
        self.button_exit = tk.Button(self,text='  Выход  ', command=self.destroy,
                                     background="#f8f8ff", relief=tk.SOLID)
        self.label2.pack()
        self.label1.pack()
        self.button_dir.pack()
        self.label6.pack()
        self.label3.pack()
        self.button_start.pack()
        self.label4.pack()
        self.button_exit.pack()
        self.label5.pack()

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
                   print(dirs)
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


if __name__ == '__main__':
    my_gui = MyGUI()
    my_gui.title('Converter to PDF')
    my_gui.mainloop()
#print('Success! Program converted', counter, 'files')
#print(input('Press ENTER to exit'))

# todo:
# 1. Сделать прогресс бар

