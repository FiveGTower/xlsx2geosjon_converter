import os
import wx
import subprocess

# Импорт необходимых функций из convert_xl2gj.py
from convert_xl2gj import get_input_files, read_excel_coordinates, read_csv_coordinates, generate_geojson

# Константы
ALLOWED_EXTENSIONS = [".xlsx", ".csv"]
CSV_FORMAT_OPTIONS = ["Номер, ДОЛ, ШИР", "Номер, ШИР, ДОЛ"]
OUTPUT_DIR = "result"

# -----------------------------------------------------------------------------
# InputListCtrl – кастомизированный список для входных файлов (wx.ListCtrl без заголовка)
# Хранит для каждого элемента: имя файла, полный путь и флаг ошибки.
# -----------------------------------------------------------------------------
class InputListCtrl(wx.ListCtrl):
    def __init__(self, parent, id=wx.ID_ANY, pos=wx.DefaultPosition, size=wx.DefaultSize):
        style = wx.LC_REPORT | wx.LC_NO_HEADER | wx.BORDER_SUNKEN  # множественный выбор работает по умолчанию
        super().__init__(parent, id, pos, size, style)
        self.InsertColumn(0, "", width=250)
        self.fileIndex = 0
        self.fileInfo = {}  # key = уникальный id, value = (filename, fullPath, failed_flag)
        self.Bind(wx.EVT_SIZE, self.OnResize)
        self.Bind(wx.EVT_KEY_DOWN, self.OnKeyDown)

    def OnResize(self, event):
        width = self.GetClientSize().width
        self.SetColumnWidth(0, width)
        event.Skip()

    def AddFile(self, filename, fullPath):
        for info in self.fileInfo.values():
            if info[1] == fullPath:
                return
        uid = self.fileIndex
        self.fileIndex += 1
        index = self.GetItemCount()
        self.InsertItem(index, filename)
        self.SetItemData(index, uid)
        self.fileInfo[uid] = (filename, fullPath, False)

    def RemoveSelected(self):
        selectedIndices = []
        index = self.GetFirstSelected()
        while index != wx.NOT_FOUND:
            selectedIndices.append(index)
            index = self.GetNextSelected(index)
        for index in sorted(selectedIndices, reverse=True):
            uid = self.GetItemData(index)
            self.DeleteItem(index)
            if uid in self.fileInfo:
                del self.fileInfo[uid]

    def GetSelectedFiles(self):
        selectedFiles = []
        index = self.GetFirstSelected()
        while index != wx.NOT_FOUND:
            uid = self.GetItemData(index)
            if uid in self.fileInfo:
                selectedFiles.append((index, self.fileInfo[uid][1]))
            index = self.GetNextSelected(index)
        return selectedFiles

    def GetAllFiles(self):
        return [info[1] for info in self.fileInfo.values()]

    def MarkAsFailed(self, fullPath):
        for index in range(self.GetItemCount()):
            uid = self.GetItemData(index)
            if uid in self.fileInfo:
                filename, path, failed = self.fileInfo[uid]
                if path == fullPath:
                    self.fileInfo[uid] = (filename, path, True)
                    self.SetItemBackgroundColour(index, wx.Colour("pink"))
                    break

    def OnKeyDown(self, event):
        key = event.GetKeyCode()
        # Delete: удалить выбранные элементы
        if key == wx.WXK_DELETE:
            self.RemoveSelected()
        # Ctrl+A: выделить все
        elif key == ord('A') and event.ControlDown():
            for i in range(self.GetItemCount()):
                self.SetItemState(i, wx.LIST_STATE_SELECTED, wx.LIST_STATE_SELECTED)
        else:
            event.Skip()

# -----------------------------------------------------------------------------
# OutputListCtrl – кастомизированный список для выходных файлов с поддержкой drag‑out через EVT_LIST_BEGIN_DRAG.
# -----------------------------------------------------------------------------
class OutputListCtrl(wx.ListCtrl):
    def __init__(self, parent, id=wx.ID_ANY, pos=wx.DefaultPosition, size=wx.DefaultSize):
        # Используем режим Report, без заголовка, с рамкой; множественный выбор разрешён по умолчанию.
        style = wx.LC_REPORT | wx.LC_NO_HEADER | wx.BORDER_SUNKEN
        super().__init__(parent, id, pos, size, style)
        self.InsertColumn(0, "", width=250)
        self.Bind(wx.EVT_SIZE, self.OnResize)
        # Используем EVT_LIST_BEGIN_DRAG для запуска операции drag‑out
        self.Bind(wx.EVT_LIST_BEGIN_DRAG, self.OnBeginDrag)
        self.Bind(wx.EVT_KEY_DOWN, self.OnKeyDown)
        self.itemData = {}  # словарь: key = индекс, value = полный путь файла

    def OnResize(self, event):
        width = self.GetClientSize().width
        self.SetColumnWidth(0, width)
        event.Skip()

    def AddFile(self, filename, fullPath):
        index = self.GetItemCount()
        self.InsertItem(index, filename)
        self.itemData[index] = fullPath

    def Clear(self):
        self.DeleteAllItems()  # Удаляем только элементы, оставляя уже созданную колонку
        self.itemData = {}

    def UpdateItems(self, items):
        # items: список кортежей (filename, fullPath)
        self.Clear()
        for filename, fullPath in items:
            self.AddFile(filename, fullPath)

    def GetClientData(self, index):
        return self.itemData.get(index, None)

    def OnBeginDrag(self, event):
        # Собираем все выбранные индексы
        selected_indices = []
        index = self.GetFirstSelected()
        while index != wx.NOT_FOUND:
            selected_indices.append(index)
            index = self.GetNextSelected(index)
        if not selected_indices:
            event.Skip()
            return

        dataObj = wx.FileDataObject()
        # Добавляем файлы по всем выбранным индексам
        for idx in selected_indices:
            full_path = self.GetClientData(idx)
            if full_path and os.path.isfile(full_path):
                dataObj.AddFile(os.path.abspath(full_path))
        dropSource = wx.DropSource(self)
        dropSource.SetData(dataObj)
        dropSource.DoDragDrop(wx.Drag_CopyOnly)
        event.Skip()

    def OnKeyDown(self, event):
        key = event.GetKeyCode()
        #delete на клавиатуре
        if key == wx.WXK_DELETE:
            # Удаляем все выбранные элементы
            selected_indices = []
            index = self.GetFirstSelected()
            while index != wx.NOT_FOUND:
                selected_indices.append(index)
                index = self.GetNextSelected(index)
            for idx in sorted(selected_indices, reverse=True):
                if idx in self.itemData:
                    del self.itemData[idx]
                self.DeleteItem(idx)
        elif key == ord('A') and event.ControlDown():
            # Выделяем все элементы на ctrl + a
            for i in range(self.GetItemCount()):
                self.SetItemState(i, wx.LIST_STATE_SELECTED, wx.LIST_STATE_SELECTED)
        else:
            event.Skip()


# -----------------------------------------------------------------------------
# Drop Target для приема файлов (фильтруются только разрешенные расширения)
# -----------------------------------------------------------------------------
class FileDropTarget(wx.FileDropTarget):
    def __init__(self, frame):
        super().__init__()
        self.frame = frame

    def OnDropFiles(self, x, y, filenames):
        valid_files = []
        for f in filenames:
            ext = os.path.splitext(f)[1].lower()
            if ext in ALLOWED_EXTENSIONS:
                valid_files.append(f)
        self.frame.OnFilesDropped(valid_files)
        return True

# -----------------------------------------------------------------------------
# Главное окно приложения
# -----------------------------------------------------------------------------
class MainFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="Geojson convert", size=(1200, 700))
        self.SetBackgroundColour("#66cc66")
        self.outputFiles = []  # Список кортежей (filename, fullPath) для выходных файлов
        self.InitOutputFiles()
        self.InitUI()
        self.Centre()

    def InitUI(self):
        mainPanel = wx.Panel(self)
        mainPanel.SetBackgroundColour("#66cc66")
        mainSizer = wx.BoxSizer(wx.VERTICAL)

        # Верхняя панель настроек
        settingsPanel = self.CreateSettingsPanel(mainPanel)
        mainSizer.Add(settingsPanel, 0, wx.EXPAND | wx.ALL, 10)

        # Основная панель с 3 блоками: вход, кнопки, выход
        contentPanel = wx.Panel(mainPanel)
        contentPanel.SetBackgroundColour("#66cc66")
        contentSizer = wx.BoxSizer(wx.HORIZONTAL)

        # Левая панель: "CSV, Excel"
        self.inputPanel = wx.Panel(contentPanel)
        self.inputPanel.SetBackgroundColour("#99ff99")
        leftSizer = wx.BoxSizer(wx.VERTICAL)
        leftLabel = wx.StaticText(self.inputPanel, label="CSV, Excel")
        leftLabel.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        leftSizer.Add(leftLabel, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 10)
        self.inputList = InputListCtrl(self.inputPanel)
        leftSizer.Add(self.inputList, 1, wx.EXPAND | wx.ALL, 10)
        self.inputPanel.SetSizer(leftSizer)
        self.inputPanel.SetDropTarget(FileDropTarget(self))
        contentSizer.Add(self.inputPanel, 1, wx.EXPAND | wx.ALL, 10)

        # Центральная панель: кнопки
        self.buttonPanel = wx.Panel(contentPanel)
        self.buttonPanel.SetBackgroundColour("#d3d3d3")
        btnSizer = wx.BoxSizer(wx.VERTICAL)
        btnSizer.AddStretchSpacer(1)
        self.processFileBtn = wx.Button(self.buttonPanel, label=">")
        self.processFileBtn.SetFont(wx.Font(16, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        btnSizer.Add(self.processFileBtn, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 10)
        self.processFileBtn.Bind(wx.EVT_BUTTON, self.OnProcessFile)
        self.processAllBtn = wx.Button(self.buttonPanel, label=">>")
        self.processAllBtn.SetFont(wx.Font(16, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        btnSizer.Add(self.processAllBtn, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 10)
        self.processAllBtn.Bind(wx.EVT_BUTTON, self.OnProcessAllFiles)
        self.openFolderBtn = wx.Button(self.buttonPanel, label="📁")
        self.openFolderBtn.SetFont(wx.Font(16, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        btnSizer.Add(self.openFolderBtn, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 10)
        self.openFolderBtn.Bind(wx.EVT_BUTTON, self.OnOpenFolder)
        self.deleteBtn = wx.Button(self.buttonPanel, label="🗑")
        self.deleteBtn.SetFont(wx.Font(16, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        btnSizer.Add(self.deleteBtn, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 10)
        self.deleteBtn.Bind(wx.EVT_BUTTON, self.OnDeleteFile)
        btnSizer.AddStretchSpacer(1)
        self.buttonPanel.SetSizer(btnSizer)
        contentSizer.Add(self.buttonPanel, 0, wx.EXPAND | wx.ALL, 10)

        # Правая панель: "Geojson"
        self.outputPanel = wx.Panel(contentPanel)
        self.outputPanel.SetBackgroundColour("#99ff99")
        rightSizer = wx.BoxSizer(wx.VERTICAL)
        rightLabel = wx.StaticText(self.outputPanel, label="Geojson")
        rightLabel.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        rightSizer.Add(rightLabel, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 10)
        self.outputList = OutputListCtrl(self.outputPanel)
        rightSizer.Add(self.outputList, 1, wx.EXPAND | wx.ALL, 10)
        self.outputPanel.SetSizer(rightSizer)
        contentSizer.Add(self.outputPanel, 1, wx.EXPAND | wx.ALL, 10)

        contentPanel.SetSizer(contentSizer)
        mainSizer.Add(contentPanel, 1, wx.EXPAND | wx.ALL, 10)
        mainPanel.SetSizer(mainSizer)
        self.UpdateOutputList()
        
    def CreateSettingsPanel(self, parent):
        panel = wx.Panel(parent)
        panel.SetBackgroundColour("#d3d3d3")
        sizer = wx.BoxSizer(wx.HORIZONTAL)
        csvLabel = wx.StaticText(panel, label="CSV формат:")
        csvLabel.SetFont(wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
        sizer.Add(csvLabel, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        self.csvFormatCombo = wx.ComboBox(panel, choices=CSV_FORMAT_OPTIONS, style=wx.CB_READONLY)
        self.csvFormatCombo.SetValue("Номер, ШИР, ДОЛ")
        sizer.Add(self.csvFormatCombo, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        self.excelAnchorCheck = wx.CheckBox(panel, label="Excel Geojson_ привязки")
        self.excelAnchorCheck.SetValue(False)
        sizer.Add(self.excelAnchorCheck, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        self.excelCycleCheck = wx.CheckBox(panel, label="Проверить цикл Excel")
        self.excelCycleCheck.SetValue(True)
        sizer.Add(self.excelCycleCheck, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        startCellLabel = wx.StaticText(panel, label="Ячейка начала:")
        startCellLabel.SetFont(wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
        sizer.Add(startCellLabel, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        self.excelStartCellText = wx.TextCtrl(panel, value="")
        sizer.Add(self.excelStartCellText, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        panel.SetSizer(sizer)
        return panel

    def OnFilesDropped(self, files):
        for f in files:
            basename = os.path.basename(f)
            self.inputList.AddFile(basename, f)
        self.UpdateOutputList()

    def OnProcessFile(self, event):
        selected = self.inputList.GetSelectedFiles()
        if not selected:
            return
        for index, file in sorted(selected, reverse=True):
            uid = self.inputList.GetItemData(index)
            if self.ProcessFile(file):
                self.inputList.DeleteItem(index)
                del self.inputList.fileInfo[uid]
            else:
                self.inputList.MarkAsFailed(file)
        self.UpdateOutputList()

    def OnProcessAllFiles(self, event):
        all_ids = list(self.inputList.fileInfo.keys())
        for uid in all_ids:
            file = self.inputList.fileInfo[uid][1]
            index = -1
            for i in range(self.inputList.GetItemCount()):
                if self.inputList.GetItemData(i) == uid:
                    index = i
                    break
            if self.ProcessFile(file):
                if index != -1:
                    self.inputList.DeleteItem(index)
                    del self.inputList.fileInfo[uid]
            else:
                self.inputList.MarkAsFailed(file)
        self.UpdateOutputList()

    def OnDeleteFile(self, event):
        # Если в входном списке есть выделенные элементы, удаляем их; иначе, удаляем выделенные в выходном списке.
        if self.inputList.GetFirstSelected() != wx.NOT_FOUND:
            self.inputList.RemoveSelected()
        elif self.outputList.GetFirstSelected() != wx.NOT_FOUND:
            selected = []
            index = self.outputList.GetFirstSelected()
            while index != wx.NOT_FOUND:
                selected.append(index)
                index = self.outputList.GetNextSelected(index)
            for index in sorted(selected, reverse=True):
                full_path = self.outputList.GetClientData(index)
                self.outputFiles = [item for item in self.outputFiles if item[1] != full_path]
            self.UpdateOutputList()

    def ProcessFile(self, file):
        ext = os.path.splitext(file)[1].lower()
        csv_format = self.csvFormatCombo.GetValue()
        if csv_format == "Номер, ДОЛ, ШИР":
            csv_order = ["n", "lon", "lat"]
        else:
            csv_order = ["n", "lat", "lon"]
        excel_anchor = self.excelAnchorCheck.GetValue()
        excel_cycle = self.excelCycleCheck.GetValue()
        start_cell = self.excelStartCellText.GetValue().strip() or None

        # Создаем папку "geojson" в директории исходного файла
        outDir = os.path.join(os.path.dirname(file), "geojson")
        os.makedirs(outDir, exist_ok=True)

        if ext == ".xlsx":
            polygon_coords, anchor_coords = read_excel_coordinates(file, start_cell=start_cell, cycle_check=excel_cycle)
            if polygon_coords is None:
                self.LogMessage("Error processing Excel file: " + file)
                return False
            generate_geojson(file, polygon_coords, anchor_coords, outDir, create_anchor=excel_anchor)
        elif ext == ".csv":
            polygon_coords, anchor_coords = read_csv_coordinates(file, csv_order)
            if polygon_coords is None:
                self.LogMessage("Error processing CSV file: " + file)
                return False
            generate_geojson(file, polygon_coords, anchor_coords, outDir)
        else:
            self.LogMessage("Unsupported file: " + file)
            return False

        outName = os.path.splitext(os.path.basename(file))[0] + ".geojson"
        fullOut = os.path.join(outDir, outName)
        if fullOut not in [fp for (_, fp) in self.outputFiles]:
            self.outputFiles.append((outName, fullOut))
        self.LogMessage("Processed file: " + file)
        return True

    def OnOpenFolder(self, event):
        # Если в входном списке выделен элемент, берем его; иначе – из выходного.
        if self.inputList.GetFirstSelected() != wx.NOT_FOUND:
            selected = self.inputList.GetSelectedFiles()
            # Берем первый из выделенных
            file = selected[0][1]
        elif self.outputList.GetFirstSelected() != wx.NOT_FOUND:
            file = self.outputList.GetClientData(self.outputList.GetFirstSelected())
        else:
            return
        try:
            subprocess.Popen(['explorer', '/select,', os.path.abspath(file)])
        except Exception as e:
            self.LogMessage("Error opening folder: " + str(e))

    def LogMessage(self, msg):
        print(msg)

    def UpdateOutputList(self):
        self.outputList.UpdateItems(self.outputFiles)

    def InitOutputFiles(self):
        self.outputFiles = []

# -----------------------------------------------------------------------------
# Запуск приложения
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    app = wx.App(False)
    frame = MainFrame()
    frame.InitOutputFiles()
    frame.Show()
    app.MainLoop()
