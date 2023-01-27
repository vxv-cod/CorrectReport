"""Утилита для сметного отдела для вставки подписей.
Особенности: работа с разными экземплярами Excel, нахождение общее количество открытых файлов"""

import os
from time import sleep
import win32com.client
import threading
from pythoncom import CoInitializeEx as pythoncomCoInitializeEx
from PyQt5 import QtCore, QtWidgets
import sys
import traceback
import ctypes

from okno_ui import Ui_Form
from  vxv_tnnc_SQL_Pyton import Sql

import imageZeroFon
from version import ver

# from rich import print
# from rich import inspect
# from prettytable import PrettyTable
# mytable = PrettyTable()
# os.system('CLS')

app = QtWidgets.QApplication(sys.argv)
Form = QtWidgets.QWidget()
ui = Ui_Form()
ui.setupUi(Form)
Form.show()

_translate = QtCore.QCoreApplication.translate
Title = 'CorrectReport v. 1.0' + str(ver)
Form.setWindowTitle(_translate("Form", Title))

'''Чистим "plainTextEdit" для отображения текста по умолчанию'''
ui.plainTextEdit.clear()

"""Если файл с сохраненным путем НЕ существует"""
savePathFile = os.getcwd() + "\savePath.ini"
if os.path.exists(savePathFile) == False:
    with open("savePath.ini", "w") as file:
        pass

'''Отслеживаем сигнал закрытия окна и сохраняем путь для подписей перед закрытием'''
def writeFail():
    with open("savePath.ini", "w") as file:
        file.write(str(ui.plainTextEdit_2.toPlainText()))
app.aboutToQuit.connect(writeFail)

'''Копируем адрес из файла после запуска программы'''
with open("savePath.ini", "r") as file:
    text = str(file.read())
    ui.plainTextEdit_2.setPlainText(f"{text}")

'''Обертка функции в потопк (декоратор)'''
def thread(my_func):
    def wrapper():
        threading.Thread(target=my_func, daemon=True).start()
    return wrapper

def colorBar(progBar, color):
    # progBar.setStyleSheet("QProgressBar::chunk {background-color: rgb(170, 170, 170); margin: 2px;}")
    progBar.setStyleSheet("QProgressBar::chunk {background-color: rgb("f"{color[0]}, {color[1]}, {color[2]}); margin: 2px;""}")

def Book():
    pythoncomCoInitializeEx(0)
    Excel = win32com.client.Dispatch("Excel.Application")
    # Excel.Visible = 0
    wb = Excel.ActiveWorkbook   # Получаем доступ к активной книге
    return wb

class Signals(QtCore.QObject):
    signal_Probar = QtCore.pyqtSignal(int)
    signal_label = QtCore.pyqtSignal(str)
    signal_err = QtCore.pyqtSignal(str)
    signal_bool = QtCore.pyqtSignal(bool)
    signal_color = QtCore.pyqtSignal(list)

    def __init__(self, parent=None):
        QtCore.QThread.__init__(self, parent)
        self.signal_Probar.connect(self.on_change_Probar,QtCore.Qt.QueuedConnection)
        self.signal_label.connect(self.on_change_label,QtCore.Qt.QueuedConnection)
        self.signal_err.connect(self.on_change_err,QtCore.Qt.QueuedConnection)
        self.signal_bool.connect(self.on_change_bool,QtCore.Qt.QueuedConnection)
        self.signal_color.connect(self.on_change_color,QtCore.Qt.QueuedConnection)

    '''Отправляем сигналы в элементы окна'''
    def on_change_Probar(self, s):
        ui.progressBar_1.setValue(s)
    def on_change_label(self, s):
        ui.label.setText(s)
    def on_change_err(self, s):
        QtWidgets.QMessageBox.information(Form, 'Excel не отвечает...', s)
    def on_change_color(self, s):
        colorBar(ui.progressBar_1, color = s)
    def on_change_bool(self, s):
        ui.pushButton.setDisabled(s)

sig = Signals()


'''
xxx = [[podName[i], podPatch[i]] for i in range(0, len(podName))]
mytable.field_names = ["Подпись", "Адрес"]
mytable.add_rows(xxx)
mytable.align = "l"
print(mytable)
'''


def Allobject():
    EnumWindows = ctypes.windll.user32.EnumWindows
    EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
    GetWindowText = ctypes.windll.user32.GetWindowTextW
    GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
    IsWindowVisible = ctypes.windll.user32.IsWindowVisible
    titles = []
    countExelList = []
    def foreach_window(hwnd, lParam):
        if IsWindowVisible(hwnd):
            length = GetWindowTextLength(hwnd)
            buff = ctypes.create_unicode_buffer(length + 1)
            GetWindowText(hwnd, buff, length + 1)
            titles.append((hwnd, buff.value))
        return True
    EnumWindows(EnumWindowsProc(foreach_window), 0)
    for i in range(len(titles)):
        if "- Excel" in  titles[i][1]:
            countExelList.append(1)
    countfail = sum(countExelList)
    return countfail


def GO():
    vedomost = []
    directoryPodpisi = str(ui.plainTextEdit_2.toPlainText())
    if "file:///" in directoryPodpisi:
        directoryPodpisi = directoryPodpisi[8:]
    if directoryPodpisi == '':
        sig.signal_err.emit(f"Не указана папка с подписями в формате *.jpg")
        return
    podName = []
    podPatch = []
    direct = os.listdir(directoryPodpisi)
    for filename in direct:
        direct = os.path.join(directoryPodpisi, filename)
        if os.path.isfile(direct) and ".jpg" in filename:
            podName.append(filename[:-4])
            podPatch.append(direct)
    podPatch = [i.replace("/", "\\") for i in podPatch]

    sig.signal_label.emit(f"Загрузка данных . . .")

    '''Поиск всех процессов EXCEL.EXE'''
    objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
    objSWbemServices = objWMIService.ConnectServer(".", "root\cimv2")
    colExcelInstances = objSWbemServices.ExecQuery(
            f"SELECT * FROM Win32_Process WHERE Name = 'EXCEL.EXE'")

    strPath = str(ui.plainTextEdit.toPlainText())
    if "file:///" in strPath:
        strPath = strPath[8:]
    if strPath == '':
        sig.signal_err.emit(f"Не указана папка для сохранения файлов")
        return
    strPath = strPath.replace("/", "\\")

    countfail = Allobject()

    nomerfail = 0
    for objInstancei in colExcelInstances:
        # objExcel = win32com.client.GetObject(None, "Excel.Application")
        objExcel = win32com.client.Dispatch("Excel.Application")
        """Отключение уведомлений с ответом по умолчанию для сохранения без подтверждения"""
        objExcel.DisplayAlerts = False      
        for objWorkbook in objExcel.Workbooks:
            nomerfail += 1
            WbName = objWorkbook.Name

            strFileExtension = ".xlsx"
            if objWorkbook.FileFormat == 52:
                strFileExtension = ".xlsm"
            if '.xls' in objWorkbook.Name:
                strFileExtension = ""
            
            wb = objWorkbook
            sheet = wb.ActiveSheet
            EndRow_1 = sheet.Cells(sheet.Rows.Count, 1).End(3).Row
            EndRow_2 = sheet.Cells(sheet.Rows.Count, 2).End(3).Row
            EndRow = max(EndRow_1, EndRow_2)


            '''
            ЛСР по Методике 2020 (БИМ)                  # 0
            Полный локальный сметный расчёт             # 1
            ОС по Методике 2020 (Приложение №5)         # 2
            Объектная смета                             # 3
            Ведомость ресурсов 8 граф с итогами БЦ      # 4
            '''            
            NameSmetList = [
                    ['ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ (СМЕТА)', 'Приложение № 2'],     # 0
                    ['ЛОКАЛЬНАЯ СМЕТА', '(наименование стройки)'],                                    # 1
                    ['ОБЪЕКТНЫЙ СМЕТНЫЙ РАСЧЕТ (СМЕТА)', 'Приложение № 5'],     # 2
                    ['ОБЪЕКТНАЯ СМЕТА', 'Форма № 3'],                           # 3
                    ['Ресурсы подрядчика', 'Трудозатраты']                      # 4
            ]
            

            xxx = sheet.Range("A1:N12").Formula
            searchList = [g  for i in xxx for g in i if g != '']
            searchList = '\n'.join(searchList)
            

            NameSmet = None
            for x in NameSmetList:
                if x[0] in searchList and x[1] in searchList:
                    NameSmet = x
            print(f"NameSmet = {NameSmet}")

            '''ЛСР по Методике 2020 (БИМ)'''
            if NameSmet == NameSmetList[0]:
                S = sheet.Range("A12").Formula
                if '№' in S:
                    TextDoc = sheet.Range("A7").Formula + sheet.Range("A10").Formula
                else:
                    S = sheet.Range("A18").Formula
                    TextDoc = sheet.Range("A13").Formula + sheet.Range("A16").Formula
                NomerDoc = S[S.rfind('№') + 1 : ].strip()

                vedomost.append([nomerfail, NomerDoc, TextDoc])
                sostavilRow = EndRow - 2
                proverilRow = EndRow
                collpod = 2


            '''Полный локальный сметный расчёт'''
            if NameSmet == NameSmetList[1]:
                sheet.Range("A1:M5").ClearContents()
                
                S = sheet.Range("D9").Formula
                NomerDoc = S[S.rfind('№') + 1 : ].strip()
                TextDoc = sheet.Range("C12").Formula

                vedomost.append([nomerfail, NomerDoc, TextDoc])
                sostavilRow = EndRow - 4
                proverilRow = EndRow - 1
                collpod = 1

            
            '''ОС по Методике 2020 (Приложение №5)'''
            if NameSmet == NameSmetList[2]:
                S = sheet.Range("B10").Formula
                NomerDoc = S[S.rfind('№') + 1 : ].strip()
                TextDoc = sheet.Range("B4").Formula + sheet.Range("B7").Formula

                vedomost.append([nomerfail, NomerDoc, TextDoc])
                proverilRow = EndRow
                sostavilRow = EndRow - 2
                NachalnikRow = EndRow - 4
                GIPRow = EndRow - 6
                collpod = None

            
            '''Объектная смета'''
            if NameSmet == NameSmetList[3]:
                # S = sheet.Range("E5").Formula
                # NomerDoc = S[S.rfind('№') + 1 : ].strip()
                NomerDoc = sheet.Range("G5").Formula
                TextDoc = sheet.Range("B2").Formula + sheet.Range("D8").Formula

                vedomost.append([nomerfail, NomerDoc, TextDoc])
                proverilRow = EndRow - 1
                sostavilRow = EndRow - 4
                NachalnikRow = EndRow - 7
                GIPRow = EndRow - 10
                collpod = 1

            '''Ведомость ресурсов 8 граф с итогами БЦ'''
            if NameSmet == NameSmetList[4]:
                S = sheet.Range("C4").Formula
                NomerDoc = S[S.rfind('№') + 1 : ].strip()
                TextDoc = sheet.Range("B1").Formula + sheet.Range("B2").Formula

                vedomost.append([nomerfail, NomerDoc, TextDoc])

            print(f"NameSmet = {NameSmet}")
            objWorkbook.SaveAs(f"{strPath}\\{objWorkbook.Name}{strFileExtension}", FileFormat=objWorkbook.FileFormat, CreateBackup=0)

            if NameSmet != NameSmetList[4]:
                
                '''Ищем в строке "S" последнее вхождение '_', забираем от следующего индекса ФИО,
                удаляем пробелы с обоих сторон, отбрасываем инициалы'''
                def poiskFamimliy(man, Row, collpod):
                    famil = None
                    try:
                        if NameSmet == NameSmetList[1]:
                            S = str(sheet.Cells(Row, collpod).Formula)
                            famil = S[S.rfind('_') + 1 :].strip()[5:]

                        if NameSmet == NameSmetList[0]:
                            S = str(sheet.Cells(Row, collpod + 1).Formula)
                            famil = S[S.find('(') + 1 : S.find(')')].strip()[5:]

                        if NameSmet == NameSmetList[2]:
                            xxx = sheet.Range(f"C{Row}:H{Row}").Formula
                            S = [g  for i in xxx for g in i if g != '']
                            S = str(S)
                            famil = S[S.find('(') + 1 : S.find(')')].strip()[5:]

                        if NameSmet == NameSmetList[3]:
                            S = str(sheet.Cells(Row, collpod).Formula)
                            famil = S[S.rfind('_') + 1 :].strip()[5:]        

                    except:                    
                        sig.signal_err.emit(f"Не найдена фамилия {man}\n{WbName}")
                    
                    return famil

                textsost = poiskFamimliy("исполнителя", sostavilRow, collpod)
                textprov = poiskFamimliy("проверяющего", proverilRow, collpod)

                if NameSmet == NameSmetList[2] or NameSmet == NameSmetList[3]:
                    textNachal = poiskFamimliy("Начальника", NachalnikRow, collpod)
                    textGIP = poiskFamimliy("ГИПа", GIPRow, collpod)

                if textsost == '':
                    sig.signal_err.emit(f"Не найдена фамилия 'Исполнителя'\n{WbName}")
                    return

                if textprov == '':
                    sig.signal_err.emit(f"Не найдена фамилия 'Проверяющего'\n{WbName}")
                    return
                
                if NameSmet == NameSmetList[2] or NameSmet == NameSmetList[3]:
                    if textNachal == '':
                        sig.signal_err.emit(f"Не найдена фамилия 'Начальника'\n{WbName}")
                        return

                    if textGIP == '':
                        sig.signal_err.emit(f"Не найдена фамилия 'ГИПа'\n{WbName}")
                        return


                sostPatch, provPatch = None, None
                for i in podPatch:
                    if textsost in i:
                        sostPatch = i
                    if textprov in i:
                        provPatch = i

                """Подписи не найдены"""
                if sostPatch == None:
                    text = f"Не найдена подпись для фамилии: {textsost}\n{WbName}"
                    sig.signal_err.emit(text)
                    return
                if provPatch == None:
                    text = f"Не найдена подпись для фамилии: {textprov}\n{WbName}"
                    sig.signal_err.emit(text)
                    return
                
                if NameSmet == NameSmetList[2] or NameSmet == NameSmetList[3]:
                    NachalPatch, GIPPatch = None, None
                    for i in podPatch:
                        if textNachal in i:
                            NachalPatch = i
                        if textGIP in i:
                            GIPPatch = i
                    
                    if NachalPatch == None:
                        text = f"Не найдена подпись для фамилии: {NachalPatch}\n{WbName}"
                        sig.signal_err.emit(text)
                        return
                    if GIPPatch == None:
                        text = f"Не найдена подпись для фамилии: {textGIP}\n{WbName}"
                        sig.signal_err.emit(text)
                        return
                
                sheet.Activate()

                for i in sheet.Shapes: i.Delete()

                def centrCell(cell_Height, img_Height):
                    centr = img_Height * 0.5 - cell_Height * 0.5
                    return  -  centr
                
                '''Полный локальный сметный расчёт'''
                if NameSmet == NameSmetList[1]:
                    col = 4
                    delta = 50
                    sheet.Cells(sostavilRow - 1, col).Activate()
                    sheet.Pictures().Insert(rf"{imageZeroFon.GO(sostPatch)}")
                    sheet.Shapes(1).IncrementLeft(delta)
                    sheet.Cells(proverilRow - 1, col).Activate()
                    sheet.Pictures().Insert(rf"{imageZeroFon.GO(provPatch)}")
                    sheet.Shapes(2).IncrementLeft(delta)
                
                '''ЛСР по Методике 2020 (БИМ)'''
                if NameSmet == NameSmetList[0]:
                    col = 5
                    cellsostav = sheet.Cells(sostavilRow, col)
                    cellsostav.Activate()
                    sheet.Pictures().Insert(rf"{imageZeroFon.GO(sostPatch)}")
                    cellprover = sheet.Cells(proverilRow, col)
                    cellprover.Activate()
                    sheet.Pictures().Insert(rf"{imageZeroFon.GO(provPatch)}")
                    
                    cell_Height = cellsostav.Height
                    img = sheet.Shapes(1)
                    img_Height = img.Height
                    img.IncrementLeft(60)
                    img.IncrementTop(centrCell(cell_Height, img_Height))

                    cell_Height = cellprover.Height
                    img = sheet.Shapes(2)
                    img.Height = img.Height
                    img.IncrementLeft(60)
                    img.IncrementTop(centrCell(cell_Height, img_Height))
                
                'ОС по Методике 2020 (Приложение №5)'
                if NameSmet == NameSmetList[2]:
                    col = 4
                    cellsostav = sheet.Cells(sostavilRow, col)
                    cellsostav.Activate()
                    sheet.Pictures().Insert(rf"{imageZeroFon.GO(sostPatch)}")

                    cellprover = sheet.Cells(proverilRow, col)
                    cellprover.Activate()
                    sheet.Pictures().Insert(rf"{imageZeroFon.GO(provPatch)}")

                    cellNachal = sheet.Cells(NachalnikRow, col)
                    cellNachal.Activate()
                    sheet.Pictures().Insert(rf"{imageZeroFon.GO(NachalPatch)}")

                    cellGIP = sheet.Cells(GIPRow, col)
                    cellGIP.Activate()
                    sheet.Pictures().Insert(rf"{imageZeroFon.GO(GIPPatch)}")
                    
                    delta = -60
                    cell_Height = cellsostav.Height
                    img = sheet.Shapes(1)
                    img_Height = img.Height
                    img.IncrementTop(centrCell(cell_Height, img_Height))
                    img.IncrementLeft(delta)

                    cell_Height = cellprover.Height
                    img = sheet.Shapes(2)
                    img.Height = img.Height
                    img.IncrementTop(centrCell(cell_Height, img_Height))
                    img.IncrementLeft(delta)

                    cell_Height = cellNachal.Height
                    img = sheet.Shapes(3)
                    img.Height = img.Height
                    img.IncrementTop(centrCell(cell_Height, img_Height))
                    img.IncrementLeft(delta)

                    cell_Height = cellGIP.Height
                    img = sheet.Shapes(4)
                    img.Height = img.Height
                    img.IncrementTop(centrCell(cell_Height, img_Height))
                    img.IncrementLeft(delta)

                '''Объектная смета'''
                if NameSmet == NameSmetList[3]:
                    col = 5
                    sheet.Cells(sostavilRow - 1, col).Activate()
                    sheet.Pictures().Insert(rf"{imageZeroFon.GO(sostPatch)}")
                    sheet.Shapes(1).IncrementLeft(-15)
                    
                    sheet.Cells(proverilRow - 1, col).Activate()
                    sheet.Pictures().Insert(rf"{imageZeroFon.GO(provPatch)}")
                    sheet.Shapes(2).IncrementLeft(-10)

                    sheet.Cells(NachalnikRow - 1, col).Activate()
                    sheet.Pictures().Insert(rf"{imageZeroFon.GO(NachalPatch)}")
                    # sheet.Shapes(3).IncrementLeft(-50)

                    sheet.Cells(GIPRow - 1, col).Activate()
                    sheet.Pictures().Insert(rf"{imageZeroFon.GO(GIPPatch)}")
                    sheet.Shapes(4).IncrementLeft(20)       

            '''Экспорт в PDF'''
            pdfName = objWorkbook.Name if ".xls" not in objWorkbook.Name else objWorkbook.Name.split(".xls")[0]
            # sheet.PrintOut(Copies=1, ActivePrinter="Microsoft Print to PDF", PrintToFile=True, PrToFileName = f"{strPath}\\{pdfName}.pdf")
            OutputFile = f"{strPath}\\{pdfName}.pdf"
            objWorkbook.ExportAsFixedFormat(0, OutputFile)
            # objWorkbook.Close(False)
            proc = round(nomerfail / countfail * 100)
            sig.signal_Probar.emit(proc)
            sig.signal_label.emit(f"{nomerfail} / {countfail}  >>  {WbName}")

        objExcel.Quit()
        objInstancei.Terminate
        sleep(2)

    sig.signal_label.emit("Формирование Ведомости смет . . .")
    Excel = win32com.client.Dispatch("Excel.Application")
    Excel.DisplayAlerts = False
    Excel.Visible=1
    wb = Excel.Workbooks.Open(os.getcwd() + "\Ведомость смет.xltx")
    sheet = wb.ActiveSheet
    EndRow = len(vedomost) + 2
    EndColl = 3
    cells = sheet.Range(sheet.Cells(3, 1), sheet.Cells(EndRow, EndColl))
    cells.Value = vedomost
    cells.Borders.Weight = 2
    
    name = wb.Name[:-1] + ".xlsx"
    try:        
        wb.SaveAs(f"{strPath}\\{name}", FileFormat=wb.FileFormat, CreateBackup=0)
    except:
        text = f"Отмена сохранения файла Ведомость смет.xltx"
        sig.signal_err.emit(text)



@thread
def start():
    Sql("CorrectReport")
    sig.signal_Probar.emit(0)
    try:
        # Sql("CorrectReport")
        sig.signal_bool.emit(True)
        sig.signal_Probar.emit(0)
        sig.signal_color.emit([100, 150, 150])
        GO()
    except:
        errortext = traceback.format_exc()
        print(errortext)
        text = f"Ошибка работы, повторите попытку \n\n{errortext}"
        sig.signal_err.emit(text)
    sig.signal_bool.emit(False)
    sig.signal_color.emit([170, 170, 170])
    sig.signal_Probar.emit(100)
    sig.signal_label.emit("Выполнено")

ui.pushButton.clicked.connect(start)
# ui.pushButton.clicked.connect(GO)

if __name__ == "__main__":
    # start()
    sys.exit(app.exec_())
    