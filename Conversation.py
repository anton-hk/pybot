import comtypes.client
from PIL import Image
from reportlab.pdfgen import canvas
from os import listdir
import os
from os.path import isfile, join


class Convert:

    #  Функция для парсинга названия файля на тело названия и формат
    def parse_name(self, name):
        doc_format = name.split('.')[-1]
        n = name.split('.')[:-1]
        name_body = ''
        for i in range(len(n)):
            name_body = name_body + str(n[i]) + '.'
        name_body = name_body[:-1]
        return (name_body, doc_format)


    # Функция для конвертации вордовских файлов в pdf
    def conversation_word(self, name):
        name_body = self.parse_name(name)[0]
        WdFormatPDF = 17
        in_file = r'C:\Users\Lumpen\PycharmProjects\pdf_bot\%(name)s' % {"name": name}
        out_file = r'C:\Users\Lumpen\PycharmProjects\pdf_bot\%(name)s.pdf' % {"name": name_body}
        word = comtypes.client.CreateObject('Word.Application')

        doc = word.Documents.Open(in_file)
        doc.SaveAs(out_file, FileFormat=WdFormatPDF)  # conversion
        doc.Close()
        word.Quit()


    #  Функция для конвертации файлов jpg в pdf
    def conversation_jpg(self, name):
        doc_format = self.parse_name(name)[1]
        name_body = self.parse_name(name)[0]
        path = r'C:\Users\Lumpen\PycharmProjects\pdf_bot\%(name)s' % {"name": name}
        pdf_path = r'C:\Users\Lumpen\PycharmProjects\pdf_bot\%(name)s.pdf' % {"name": name_body}
        image1 = Image.open(path)
        if doc_format == 'png':
            rgb = Image.new('RGB', image1.size, (255, 255, 255))
            rgb.paste(image1)
            rgb.save(pdf_path, 'PDF', resolution=100.0)
            rgb.close()
        if doc_format == 'jpg':
            image1.convert('RGB')
            image1.save(pdf_path)
        image1.close()

    # Фунция для конвертации списка картинок в папке photos
    def conversation_list_images(self):
        path = r'C:\Users\Lumpen\PycharmProjects\pdf_bot\photos'
        pdf_path = r'C:\Users\Lumpen\PycharmProjects\pdf_bot\photos\file.pdf'
        im1 = Image.open(os.path.join(path, listdir(path)[0]))
        lim = []
        for each_file in listdir(path):
            image_path = os.path.join(path, each_file)
            img = Image.open(image_path)
            img.convert('RGB')
            lim.append(img)
        lim.remove(lim[0])
        im1.save(pdf_path, save_all=True, append_images=lim)