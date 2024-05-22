from PyQt6.QtCore import *
from PyQt6.QtWidgets import *
import openpyxl
import sys
import time

# Изменение метода __lt__ чтобы можно было производить поиск текстом
class TableWidgetItem(QTableWidgetItem):
    def __lt__(self, other):
        try:
            return float(self.text()) < float(other.text())
        except ValueError:
            return super().__lt__(other)

# Основное окно приложения
class Window(QMainWindow):
    def __init__(self):
        super().__init__()

        self.menu_bar = QMenuBar()

        self.menu = QMenu("Файлик")
        action1 = self.menu.addAction("Сохранить Файл")
        action1.triggered.connect(self.export)

        self.menu_bar.addMenu(self.menu)

        self.setMenuBar(self.menu_bar)

    def closeEvent(self, event):
        widget.DataWindow.close()

        event.accept()

    # Сохранение файла
    def export(self):
        filename, filter = QFileDialog.getSaveFileName(self, 'Save file', '','Excel files (*.xlsx)')
        wb = openpyxl.Workbook()

        temp = wb["Sheet"]
        wb.remove(temp)

        print(filename)

        centralWidget = self.centralWidget()

        #Запись информации в новый файл
        for sheet, table_widget in centralWidget.temp_sheets.items():
            workingSheet = wb.create_sheet(sheet.title)

            labels = []
            for c in range(table_widget.columnCount()):
                it = table_widget.horizontalHeaderItem(c)
                labels.append(str(c+1) if it is None else it.text())
            print(labels)

            for i in range(len(labels)):
                value = labels[i]
                workingSheet.cell(1, i + 1, value)

            if filename:
                for column in range(table_widget.columnCount()):
                    for row in range(table_widget.rowCount()):
                        try:
                            text = str(table_widget.item(row, column).text())
                            workingSheet.cell(row + 2, column + 1, text)
                        except AttributeError:
                            pass
                        
        wb.save(filename)

# Окно для добавления новой даты
class DataWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Добавить Дату")
        self.setFixedSize(500, 500)

        self.Layout = QVBoxLayout()
        self.setLayout(self.Layout)

        self.addButton = QPushButton("Добавить")
        self.addButton.clicked.connect(self.acceptInfo)
        self.Layout.addWidget(self.addButton)
        
        self.inputs :list[QLineEdit] = []
        self.labels :list[QLabel] = []

    # Добавления новой строки по записанным данным
    def acceptInfo(self):
        widget.currList.insertRow(0)

        for i in range(len(self.labels)):
            print(0, i)
            item = TableWidgetItem(str(self.inputs[i].text()))

            widget.currList.setItem(0 , i, item)

    #Загрузка полей для заполнений
    def load(self):
        self.Layout.removeWidget(self.addButton)

        labels = []

        for c in range(widget.currList.columnCount()):
            label = widget.currList.horizontalHeaderItem(c)
            labels.append(str(c+1) if label is None else label.text())
        
        for l in range(len(labels)):
            label = QLabel(labels[l])
            lineEdit = QLineEdit()
            self.Layout.addWidget(label)
            self.Layout.addWidget(lineEdit)

            self.inputs.append(lineEdit)
            self.labels.append(label)

        self.Layout.addWidget(self.addButton)

    # Действия при закрытии окна
    def closeEvent(self, event):
        for i in reversed(range(self.Layout.count())): 
            self.Layout.itemAt(i).widget().setParent(None)

        event.accept()

# Главное окно где создаются списки и редактируется дата
class Main(QWidget):
    def __init__(self):
        super(Main, self).__init__()
        self.setWindowTitle("Курсовая")
        self.setFixedSize(900, 900)
        
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        self.addData = QPushButton("Добавить")
        self.remove_row = QPushButton("Удалить")
    
        self.filter = QComboBox()

        self.DataWindow = DataWindow()

        self.addData.clicked.connect(self.show_data_window)
        self.remove_row.clicked.connect(self.deleteRow)

        self.selectedRow = 0

        self.search = QLineEdit()
        self.search.setPlaceholderText("Ввести поиск здесь...")
        self.search.textChanged.connect(self.findName)

        self.list = QListWidget()
        self.list.setFixedSize(900, 70)
        self.list.itemClicked.connect(self.onClickList)

        #Добавление виджетов pyqt6 в макет/каркас
        layout.addWidget(self.list)
        layout.setAlignment(self.list, Qt.AlignmentFlag.AlignCenter)

        layout.addWidget(self.filter)
        layout.addWidget(self.search)
        layout.addWidget(self.addData)
        layout.addWidget(self.remove_row)
        
        path = "./data.xlsx"
        self.workbook = openpyxl.load_workbook(path)
        self.temp_sheets = {}

        # Предзагрузка листов
        for name in self.workbook.sheetnames:
            tableWidget = QTableWidget()

            tableWidget.cellClicked.connect(self.cellClicked)

            self.temp_sheets[self.workbook[name]] = tableWidget
            tableWidget.hide()

            layout.addWidget(tableWidget)

        self.currList :QTableWidget = self.temp_sheets[next(iter(self.temp_sheets))]

        self.load_data()

    #Запуск окна с добавлением инфы
    def show_data_window(self):
        self.DataWindow.show()
        self.DataWindow.load()

    #Сортировка
    def findName(self):
        name = self.search.text().lower()
        found = False

        #Поиск по заданной строке
        for row in range(self.currList.rowCount()):
            for column in range(self.currList.columnCount()):
                header = self.currList.horizontalHeaderItem(column)
                if header != None:
                    if header.text() == self.filter.currentText(): 
                        item = self.currList.item(row, column)
                        try:
                            found = name in item.text().lower()
                            self.currList.setRowHidden(row, not found)
                            if found:
                                break
                        except AttributeError:
                            pass

    #Переключение между листами Excel при помощи QListWidget
    def onClickList(self, item :QListWidgetItem):
        for sheet, widget in self.temp_sheets.items():
            widget.hide()

        self.temp_sheets[self.workbook[item.text()]].show()
        self.currList = self.temp_sheets[self.workbook[item.text()]]

        labels = []

        for c in range(self.currList.columnCount()):
            label = self.currList.horizontalHeaderItem(c)
            labels.append(str(c+1) if label is None else label.text())
    
        #Загрузка фильтра при смене листа
        self.filter.clear()

        for i in labels:
            self.filter.addItem(i)

    #Загрузка всех листов Excel при помощи .loadSheet()
    def load_data(self):

        for name in self.workbook.sheetnames:
            self.list.addItem(name)
            self.loadSheet(self.workbook[name])

    #Загрузить лист Excel
    def loadSheet(self, sheet):

        self.temp_sheets[sheet].setRowCount(sheet.max_row)
        self.temp_sheets[sheet].setColumnCount(sheet.max_column)
        
        list_values = list(sheet.values)
        self.temp_sheets[sheet].setHorizontalHeaderLabels(list_values[0])
        self.temp_sheets[sheet].setSortingEnabled(True)

        row_index = 0

        # Добавление информации из листа Excel в QTableWidget
        for value_tuple in list_values[1:]:
            col_index = 0
            for value in value_tuple:
                item = TableWidgetItem(str(value))
                
                self.temp_sheets[sheet].setItem(row_index , col_index, item)
               
                col_index += 1
            row_index += 1

    #Добавить линию
    def addRow(self):
        self.currList.insertRow(self.selectedRow)

        return self.selectedRow

    #Удалить линию
    def deleteRow(self, *args):
        print(self.selectedRow)
        self.currList.removeRow(self.selectedRow)
        print(*args)
    
    # Сохранить текущую клетку
    def cellClicked(self, row, column):
        self.selectedRow = row
        

#Запуск кода
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Window()
    widget = Main()

    window.setCentralWidget(widget)
    widget.show()
    window.show()
    
    app.exec()