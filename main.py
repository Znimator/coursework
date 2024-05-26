from PyQt6.QtCore import *
from PyQt6.QtWidgets import *
from PyQt6.QtMultimedia import *
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

        self.menu = QMenu("Файл")
        action1 = self.menu.addAction("Сохранить")
        action1.triggered.connect(self.export)
        self.setStyleSheet("background-color: #ADADAD; color: black;")

        self.menu_bar.addMenu(self.menu)

        self.setMenuBar(self.menu_bar)

    def closeEvent(self, event):
        print("closing")
        widget.DataWindow.close()

        event.accept()

    def export(self):
        filename, filter = QFileDialog.getSaveFileName(self, 'Save file', '','Excel files (*.xlsx)')
        wb = openpyxl.Workbook()

        temp = wb["Sheet"]
        wb.remove(temp)

        print(filename)

        centralWidget = self.centralWidget()

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

        self.setWindowTitle("Add Data")
        #self.setFixedSize(500, 500)
        self.setStyleSheet("background-color: #ADADAD; color: black;")

        self.Layout = QVBoxLayout()
        self.setLayout(self.Layout)

        self.addButton = QPushButton("Добавить")
        self.cancelButton = QPushButton("Отмена")
        self.cancelButton.clicked.connect(self.cancel)
        self.addButton.clicked.connect(self.acceptInfo)
        self.Layout.addWidget(self.addButton)
        
        self.inputs :list[QLineEdit] = []

    def cancel(self):
        self.close()

    # Добавления новой строки по записанным данным
    def acceptInfo(self):
        widget.currentWidget.insertRow(0)

        for i in range(len(self.inputs)):
            print(0, i)
            item = TableWidgetItem(str(self.inputs[i].text()))

            widget.currentWidget.setItem(0 , i, item)

    #Загрузка полей для заполнений
    def load(self):
        print("loading stuff")

        self.Layout.removeWidget(self.addButton)

        labels = []

        for c in range(widget.currentWidget.columnCount()):
            label = widget.currentWidget.horizontalHeaderItem(c)
            labels.append(str(c+1) if label is None else label.text())
        
        for l in range(len(labels)):
            lineEdit = QLineEdit()
            lineEdit.setPlaceholderText(labels[l])
            self.Layout.addWidget(lineEdit)

            self.inputs.append(lineEdit)

        self.Layout.addWidget(self.addButton)

    # Действия при закрытии окна
    def closeEvent(self, event):
        print("clearing stuff")

        for i in reversed(range(self.Layout.count())): 
            self.Layout.itemAt(i).widget().setParent(None)
            
        self.inputs :list[QLineEdit] = []

        event.accept()

# Главное окно где создаются списки и редактируется дата
class Main(QWidget):
    def __init__(self):
        super(Main, self).__init__()
        self.setWindowTitle("Load Excel data to QTableWidget")
        self.setBaseSize(700, 700)
        
        self.main_layout = QVBoxLayout()
        self.setLayout(self.main_layout)
        
        self.addData = QPushButton("Добавить Информацию")
        self.add_row = QPushButton("Добавить линию")
        self.remove_row = QPushButton("Удалить линию")
    
        self.music = QSoundEffect()
        self.music.setSource(QUrl.fromLocalFile("./music.wav"))
        self.music.setLoopCount(10)
        self.music.setVolume(0.25)
        self.music.play()

        self.player = QSoundEffect()
        self.player.setSource(QUrl.fromLocalFile("./click.wav"))
        self.player.setVolume(0.75)

        self.filter_list = QComboBox()
        self.filter_list.setPlaceholderText("Фильтры")
        self.sheet_list = QComboBox()
        self.sheet_list.setPlaceholderText("Листы")
        self.sheet_list.currentTextChanged.connect(self.listChange)

        self.DataWindow = DataWindow()

        self.addData.clicked.connect(self.show_data_window)

        self.add_row.clicked.connect(self.addRow)
        self.remove_row.clicked.connect(self.deleteRow)

        self.selectedRow = 0

        self.search = QLineEdit()
        self.search.textChanged.connect(self.findName)

        #Добавление виджетов pyqt6 в макет/каркас
        self.main_layout.addWidget(self.filter_list)
        self.main_layout.addWidget(self.sheet_list)
        self.main_layout.addWidget(self.search)
        self.main_layout.addWidget(self.addData)
        self.main_layout.addWidget(self.remove_row)

        
        path = "./data.xlsx"
        self.workbook = openpyxl.load_workbook(path)
        self.temp_sheets = {}

        self.buttons :dict[str, QPushButton] = {}

        # Предзагрузка листов
        for name in self.workbook.sheetnames:
            tableWidget = QTableWidget()

            tableWidget.cellClicked.connect(self.cellClicked)
            self.temp_sheets[self.workbook[name]] = tableWidget
            self.sheet_list.addItem(name)
            tableWidget.hide()

            self.main_layout.addWidget(tableWidget)

        self.currentWidget :QTableWidget = self.temp_sheets[next(iter(self.temp_sheets))]

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
        for row in range(self.currentWidget.rowCount()):
            for column in range(self.currentWidget.columnCount()):
                header = self.currentWidget.horizontalHeaderItem(column)
                if header != None:
                    if header.text() == self.filter_list.currentText(): 
                        item = self.currentWidget.item(row, column)
                        try:
                            found = name in item.text().lower()
                            self.currentWidget.setRowHidden(row, not found)
                            if found:
                                break
                        except AttributeError:
                            pass

    #Переключение между листами Excel при помощи QListWidget
    def listChange(self, item):
        self.player.play()

        for sheet, widget in self.temp_sheets.items():
            widget.hide()

        self.temp_sheets[self.workbook[item]].show()
        self.currentWidget = self.temp_sheets[self.workbook[item]]

        labels = []

        for c in range(self.currentWidget.columnCount()):
            label = self.currentWidget.horizontalHeaderItem(c)
            labels.append(str(c+1) if label is None else label.text())
    
        #Загрузка фильтра при смене листа
        self.filter_list.clear()

        for i in labels:
            self.filter_list.addItem(i)

    #Загрузка всех листов Excel при помощи .loadSheet()
    def load_data(self):

        for name in self.workbook.sheetnames:
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
        self.currentWidget.insertRow(self.selectedRow)

        return self.selectedRow

    #Удалить линию
    def deleteRow(self, *args):
        self.player.play()

        self.currentWidget.removeRow(self.selectedRow)
        
    
    # Сохранить текущую клетку
    def cellClicked(self, row, column):
        self.player.play()
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