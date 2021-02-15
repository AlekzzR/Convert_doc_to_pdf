import os
import comtypes.client
#4
wdFormatPDF = 17  # number of format in M.Word application

input_dir = r'C:\Users\alekz\OneDrive\Рабочий стол\Загрузка ФГИС\Февраль\Бывалово\Январь\docx'
output_dir = r'C:\Users\alekz\OneDrive\Рабочий стол\Загрузка ФГИС\Февраль\Бывалово\Январь\pdf'
try:
    os.stat(output_dir)  # check specified directory
except:
    os.mkdir(output_dir)  # create a directory if dosnt have
for root, dirs, files in os.walk(input_dir):
    for file in files:
        in_file = os.path.join(root, file)
        print(file)
        output_file = file.split('.')[0]  # splits the str by 2, makes a list, and passes the first str to the variable
        out_file = output_dir + '\\' + output_file + '.pdf'
        print(out_file)
        word = comtypes.client.CreateObject('Word.Application', dynamic=True)  # without "dynamic" dont work
        word.Visible = False  # works together with both True and False
        doc = word.Documents.Open(in_file)
        doc.SaveAs(out_file, wdFormatPDF)
        doc.Close()
        word.Quit()
