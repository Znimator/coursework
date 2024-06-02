from PyQt6.QtCore import *
from PyQt6.QtWidgets import *
from PyQt6 import uic
import openpyxl
from datetime import date
import re

# Изменение метода __lt__ чтобы можно было производить поиск текстом
class TableWidgetItem(QTableWidgetItem):
    def __lt__(self, other):
        try:
            return float(self.text()) < float(other.text())
        except ValueError:
            return super().__lt__(other)
        
class InfoWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.TextLabel = QLabel()
        self._layount = QVBoxLayout()

        self._layount.addWidget(self.TextLabel)

        self.setLayout(self._layount)

    def ShowText(self, text):
        self.show()
        self.TextLabel.setText(text)

# Окно для добавления новой даты
class AddClientWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Add Client")

        self.Layout = QVBoxLayout()
        self.setLayout(self.Layout)

        self.addButton = QPushButton("Добавить")
        self.cancelButton = QPushButton("Отмена")
        self.cancelButton.clicked.connect(self.cancel)
        self.addButton.clicked.connect(self.acceptInfo)
        self.Layout.addWidget(self.addButton)
        self.Layout.addWidget(self.cancelButton)
        
        self.inputs :list[QLineEdit] = []

    def cancel(self):
        self.close()

    # Добавления новой строки по записанным данным
    def acceptInfo(self):
        sheet = sheets[0]

        rows = list(sheet.rows)

        for i in range(len(self.inputs)):
            value = self.inputs[i].text()

            sheet.cell(len(rows) + 1, i+1, value)
    
    #Загрузка полей для заполнений
    def load(self):
        print("loading stuff")

        self.Layout.removeWidget(self.addButton)
        self.Layout.removeWidget(self.cancelButton)

        labels = list(sheets[0].values)
        
        print(labels)

        for l in labels[0]:
            print(l)

            lineEdit = QLineEdit()
            lineEdit.setPlaceholderText(l)
            self.Layout.addWidget(lineEdit)

            self.inputs.append(lineEdit)

        self.Layout.addWidget(self.addButton)
        self.Layout.addWidget(self.cancelButton)

    # Действия при закрытии окна
    def closeEvent(self, event):
        print("clearing stuff")

        for i in reversed(range(self.Layout.count())): 
            self.Layout.itemAt(i).widget().setParent(None)

        self.inputs :list[QLineEdit] = []

        event.accept()

def loadSheet(index):
    tableWidget.reset()

    sheet = sheets[index]
    global currentSheetName
    currentSheetName = workbook.sheetnames[index]

    tableWidget.setRowCount(sheet.max_row - 1)
    tableWidget.setColumnCount(sheet.max_column)
    
    list_values = list(sheet.values)
    tableWidget.setHorizontalHeaderLabels(list_values[0])
    tableWidget.setSortingEnabled(True)

  #  Добавление информации из листа Excel в QTableWidget

    row_index = 0
    for value_tuple in list_values[1:]:
        col_index = 0
        
        for value in value_tuple:
            item = TableWidgetItem(str(value))
            
            tableWidget.setItem(row_index , col_index, item)
            
            col_index += 1
        row_index += 1

def loadClients():
    loadSheet(0)

def loadService():
    loadSheet(1)

def loadDone():
    loadSheet(2)

def programInfo():
    infoWindow.ShowText("ОБ ПРОГРАММЕ!!!")

def authorInfo():
    infoWindow.ShowText("ОБ АВТОРЕ!!!!!")

def addClient():
    addClientWindow.load()
    addClientWindow.show()

def finishService():

    if currentSheetName != "Клиенты":
        return

    clientSheet = sheets[0]
    serviceSheet = sheets[1]
    doneSheet = sheets[2]
    selectedLine = tableWidget.currentRow() + 1

    if selectedLine <= 0:
        return

    client = clientSheet[selectedLine + 1][0].value
    service = clientSheet[selectedLine + 1][2].value
    price = 0

    if client == None:
        return

    for i in list(serviceSheet.values)[1:]:
        if i[0] == service:
            price = i[1]

    currentDate = date.today()

    nextRow = len(list(doneSheet.rows)) + 1

    print(service, client, currentDate, price)

    doneSheet.cell(nextRow, 1, service)
    doneSheet.cell(nextRow, 2, client)
    doneSheet.cell(nextRow, 3, currentDate)
    doneSheet.cell(nextRow, 4, price)

    clientSheet.delete_rows(selectedLine + 1)

    earn = 0

    for i in range(1, len(list(doneSheet.rows))):
        earn += int(doneSheet[i+1][3].value)
        earnedLabel.setText(f"Заработанно: {earn}")
        print(doneSheet[i+1][3])
    
def export():
    filename, filter = QFileDialog.getSaveFileName(window, 'Save file', '','Excel files (*.xlsx)')
    
    workbook.save(filename)
    

Form, Window = uic.loadUiType("nasty.ui")
app = QApplication([])
window :QMainWindow = Window()
addClientWindow :QWidget = AddClientWindow()
infoWindow :InfoWindow = InfoWindow()

form = Form()
form.setupUi(window)

path = "./data.xlsx"
workbook = openpyxl.load_workbook(path)

sheets = workbook.worksheets

earnedLabel :QLabel = form.label

global currentSheetName  
currentSheetName = workbook.sheetnames[0]
print(currentSheetName)
print(sheets)

tableWidget :QTableWidget = form.tableWidget

form.pushButton.clicked.connect(loadClients)
form.pushButton_2.clicked.connect(loadService)
form.pushButton_3.clicked.connect(loadDone)
form.pushButton_5.clicked.connect(addClient)
form.pushButton_4.clicked.connect(finishService)

form.action_3.triggered.connect(export)
form.action.triggered.connect(authorInfo)
form.action_2.triggered.connect(programInfo)

earned = 0
done = sheets[2]

print(len(list(done.rows)))

for i in range(1, len(list(done.rows))):
    print("Aaaa")
    earned += int(done[i+1][3].value)
    earnedLabel.setText(f"Заработанно: {earned}")
    print(done[i+1][3])

window.show()
app.exec()