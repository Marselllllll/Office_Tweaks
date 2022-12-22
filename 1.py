import os
from tkinter import Tk
from tkinter.filedialog import askdirectory, askopenfilename, askopenfilenames
from pdf2docx import Converter
from docx2pdf import convert
from PIL import Image


def work_catalog():
    choose = int(input("\nВы хотите ввести абсолютный путь:\n"
                       "1. \033[32mВручную\033[0m\n"
                       "2. С помощью \033[32mдиалогового окна\033[0m\n"
                       "Выбор: "))
    if choose == 1:
        filename = input("\nВведите абсолютный путь: ")
        print()
        os.chdir(filename)
    else:
        Tk().withdraw()
        filename = askdirectory()
        os.chdir(filename)


def PDFtoDOCX():
    choose = int(input("\nВы хотите выбрать файл формата .pdf:\n"
                       "1. Через \033[32mконсоль\033[0m\n"
                       "2. С помощью \033[32mдиалогового окна\033[0m\n"
                       "Выбор: "))
    if choose == 1:
        lst0 = os.listdir(os.getcwd())
        if len(lst0) != 0:
            lst = []
            for i in lst0:
                if i.endswith("pdf") == True:
                    lst.append(i)
            print("\nСписок файлов с расширением .pdf в данном каталоге:\n")
            count = 1
            for k in lst:
                print(f"{count}. {k}")
                count += 1
            choose_file = int(input("Введите номер файла для преобразования (или '0' чтоб преобразовать все): "))
            if choose_file == 0:
                for i in lst:
                    cv = Converter(i)
                    cv.convert(i[:-3]+"docx")
                    cv.close()
            else:
                print(lst)
                print(count)
                cv = Converter(lst[choose_file-1])
                cv.convert(lst[choose_file-1][:-3]+"docx")
                cv.close()

    else:
        choose_method = int(input("\nВы хотите конвертировать:\n"
                                  "1. \033[32mОдин\033]0m файл\n"
                                  "2. \033[32mНесколько\033]0m файлов\n"
                                  "Выбор: "))
        if choose_method == 1:
            while True:
                Tk().withdraw()
                filename = askopenfilename()
                if filename.endswith("pdf") == True:
                    cv = Converter(filename)
                    cv.convert(filename[:-3]+"docx")
                    cv.close()
                    break
                else:
                    choose_way = int(input("\n\033[33mНеверный формат файла\033[0m\n"
                          "Повторить поппытку?"
                          "1. Да"
                          "2. Нет"
                          "Выбор: "))
                    if choose_way == 2:
                        print("\nКонвертирование не удалось\n")
                        break
        else:
            while True:
                Tk().withdraw()
                filename = askopenfilenames()
                for i in filename:
                    if i.endswith("pdf") == True:
                        check = 0
                    else:
                        check = 1

                    if check == 0:
                        for i in filename:
                            cv = Converter(i)
                            cv.convert(i[:-3] + "docx")
                            cv.close()
                            break
                    else:
                        choose_way = int(input("\n\033[33mНе все файлы формата .pdf\033[0m\n"
                                               "Повторить поппытку?"
                                               "1. Да"
                                               "2. Нет"
                                               "Выбор: "))
                        if choose_way == 2:
                            print("\nКонвертирование не удалось\n")
                            break


def DOCXtoPDF():
    choose = int(input("\nВы хотите выбрать файл формата .docx:\n"
                       "1. Через \033[32mконсоль\033[0m\n"
                       "2. С помощью \033[32mдиалогового окна\033[0m\n"
                       "Выбор: "))
    if choose == 1:
        lst0 = os.listdir(os.getcwd())
        if len(lst0) != 0:
            lst = []
            for i in lst0:
                if i.endswith("docx") == True:
                    lst.append(i)
            print("\nСписок файлов с расширением .docx в данном каталоге:\n")
            count = 1
            for k in lst:
                print(f"{count}. {k}")
                count += 1
            choose_file = int(input("Введите номер файла для преобразования (или '0' чтоб преобразовать все): "))
            if choose_file == 0:
                for i in lst:
                    convert(i, i[:-4] + "pdf")
            else:
                convert(lst[choose_file - 1], lst[choose_file - 1][:-4] + "pdf")

    else:
        choose_method = int(input("\nВы хотите конвертировать:\n"
                                  "1. \033[32mОдин\033]0m файл\n"
                                  "2. \033[32mНесколько\033]0m файлов\n"
                                  "Выбор: "))
        if choose_method == 1:
            while True:
                Tk().withdraw()
                filename = askopenfilename()
                if filename.endswith("docx") == True:
                    convert(filename, filename[:-4] + "pdf")
                    break
                else:
                    choose_way = int(input("\n\033[33mНеверный формат файла\033[0m\n"
                                           "Повторить поппытку?\n"
                                           "1. Да\n"
                                           "2. Нет\n"
                                           "Выбор: "))
                    if choose_way == 2:
                        print("\nКонвертирование не удалось\n")
                        break
        else:
            while True:
                Tk().withdraw()
                filename = askopenfilenames()
                for i in filename:
                    if i.endswith("docx") == True:
                        check = 0
                    else:
                        check = 1

                    if check == 0:
                        for i in filename:
                            convert(i, i[:-4] + "pdf")
                            break
                    else:
                        choose_way = int(input("\n\033[33mНе все файлы формата .pdf\033[0m\n"
                                               "Повторить поппытку?\n"
                                               "1. Да\n"
                                               "2. Нет\n"
                                               "Выбор: "))
                        if choose_way == 2:
                            print("\nКонвертирование не удалось\n")
                            break


def image():
    choose = int(input("\nВы хотите выбрать файл формата .gif, .png, .jpg, .jpeg:\n"
                       "1. Через \033[32mконсоль\033[0m\n"
                       "2. С помощью \033[32mдиалогового окна\033[0m\n"
                       "Выбор: "))
    if choose == 1:
        lst0 = os.listdir(os.getcwd())
        if len(lst0) != 0:
            lst = []
            for i in lst0:
                if i.endswith("gif") or i.endswith("png") or i.endswith("jpg") or i.endswith("jpeg") == True:
                    lst.append(i)
            print("\nСписок файлов с расширением .docx в данном каталоге:\n")
            count = 1
            for k in lst:
                print(f"{count}. {k}")
                count += 1
            choose_file = int(input("Введите номер файла для преобразования (или '0' чтоб преобразовать все): "))
            if choose_file == 0:
                choose_increase = int(input("\nВведите параметр сжатия (от '0' до '100'\n"
                                            "Ввод: "))
                for i in lst:
                    image_file = Image.open(i)
                    image_file.save(i, quality=choose_increase)
            else:
                choose_increase = int(input("\nВведите параметр сжатия (от '0' до '100')\n"
                                            "Ввод: "))
                image_file = Image.open(lst[choose_file-1])
                image_file.save(lst[choose_file-1], quality=choose_increase)
    else:
        choose_method = int(input("\nВы хотите сжать:\n"
                                  "1. \033[32mОдин\033]0m файл\n"
                                  "2. \033[32mНесколько\033]0m файлов\n"
                                  "Выбор: "))
        if choose_method == 1:
            while True:
                Tk().withdraw()
                filename = askopenfilename()
                if filename.endswith("gif") or filename.endswith("png") or filename.endswith("jpg")\
                        or filename.endswith("jpeg") == True:
                    choose_increase = int(input("\nВведите параметр сжатия (от '0' до '100')\n"
                                                "Ввод: "))
                    image_file = Image.open(filename)
                    image_file.save(filename, quality=choose_increase)
                    break
                else:
                    choose_way = int(input("\n\033[33mНеверный формат файла\033[0m\n"
                                           "Повторить поппытку?\n"
                                           "1. Да\n"
                                           "2. Нет\n"
                                           "Выбор: "))
                    if choose_way == 2:
                        print("\nСжатие не удалось\n")
                        break
        else:
            while True:
                Tk().withdraw()
                filename = askopenfilenames()
                for i in filename:
                    if i.endswith("gif") or i.endswith("png") or i.endswith("jpg") or i.endswith("jpeg") == True:
                        check = 0
                    else:
                        check = 1

                    if check == 0:
                        choose_increase = int(input("\nВведите параметр сжатия (от '0' до '100')\n"
                                                    "Ввод: "))
                        for i in filename:
                            image_file = Image.open(i)
                            image_file.save(i, quality=choose_increase)
                            break
                    else:
                        choose_way = int(input("\n\033[33mНе все файлы формата .pdf\033[0m\n"
                                               "Повторить поппытку?\n"
                                               "1. Да\n"
                                               "2. Нет\n"
                                               "Выбор: "))
                        if choose_way == 2:
                            print("\nКонвертирование не удалось\n")
                            break
    print()

def deletefile():
    lst0 = os.listdir(os.getcwd())
    choose = int(input("Веберите действие:\n"
         "1. Удалить все файлы, начинающиеся на определённую подстроку\n"
         "2. Удалить все файлы, заканчивающиеся на определённую подстроку\n"
         "3. Удалить все файлы, содержащие определенную подстроку\n"
         "4. Удалить файлы по расширению\n"
         "Выбор: "))
    if choose == 1:
        marker = input("Введите подстроку начала: ")
        for i in lst0:
            if i.startswith(marker) == True:
                os.remove(i)
    if choose == 2:
        marker = input("Введите подстроку конца: ")
        for i in lst0:
            if i.startswith(marker) == True:
                os.remove(i)
    if choose == 3:
        marker = input("Введите подстроку: ")
        for i in lst0:
            if i.find(marker) != -1:
                os.remove(i)
    if choose == 4:
        marker = input("Введите расширение: ")
        for i in lst0:
            if i[i.rindex(".")+1:].lower() == marker.lower():
                os.remove(i)
    print()


choose = 0
while choose != 5:
    choose = int(input(f"Выберите дейтсвие:\n"
                       f"0. Сменить рабочий каталог\n"
                       f"1. Преобразовать PDF в Docx\n"
                       f"2. Преобразовать Docx в PDF\n"
                       f"3. Произвести сжатие изображений\n"
                       f"4. Удалить группу файлов\n"
                       f"5. Выход\n"
                       f"Выбор: "))
    if choose == 0:
        work_catalog()
        print(f"\nТекущий рабочий каталог: {os.getcwd()}\n")
    if choose == 1:
        PDFtoDOCX()
    if choose == 2:
        DOCXtoPDF()
    if choose == 3:
        image()
    if choose == 4:
        deletefile()
    if choose == 5:
        break

