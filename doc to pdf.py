import sys
import os, os.path
import comtypes.client

wdFormatPDF = 17

input_dir = r'C:\Users\alekz\OneDrive\Рабочий стол\Загрузка ФГИС\Февраль\Бывалово\Январь\docx'
output_dir = r'C:\Users\alekz\OneDrive\Рабочий стол\Загрузка ФГИС\Февраль\Бывалово\Январь\pdf'
os.mkdir(output_dir)
for root, dirs, files in os.walk(input_dir):
    for file in files:
        in_file = os.path.join(root, file)
        print(file)
        output_file = file.split('.')[0]
        print(output_file)
        out_file = output_dir + '\\' + output_file + '.pdf'
        print(out_file)
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = True
        doc = word.Documents.Open(in_file)
        doc.SaveAs(out_file, wdFormatPDF)
        doc.Close()
        word.Quit()
        break
