import win32com.client, datetime
from PySide2.QtCore import QFile
from PySide2.QtGui import QStandardItemModel, QStandardItem, QIcon

from PySide2.QtWidgets import QWidget, QApplication, QTableView, QStackedWidget, QMainWindow, QSizePolicy, \
    QVBoxLayout, QPushButton, QSpinBox, QHBoxLayout, QMenuBar, QAction
import sys


class MainWindowUi(QMainWindow):
    def __init__(self):
        super(MainWindowUi, self).__init__()
        self.setWindowIcon(QIcon("./data/calendar.ico"))

        with open("./data/dark.css", "r") as file:
            stylesheed = " ".join(file.readlines())
            self.setStyleSheet(stylesheed)
            print(stylesheed)

        self.setWindowTitle("checker")
        self.resize(400, 300)
        self.setup_ui()
        self.setup_menubar()
        self.setup_signals()
        self.check_updates()
        self.theme = "dark"
        self.show()

    def setup_ui(self):
        self.stacked_widget = QStackedWidget()
        self.setCentralWidget(self.stacked_widget)

        self.p_view = QWidget()
        self.stacked_widget.addWidget(self.p_view)

        self.model = QStandardItemModel(self.p_view)
        self.table = QTableView(self.p_view)
        self.table.setModel(self.model)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.table.resizeColumnToContents(1)
        self.model.setHorizontalHeaderLabels(["start", "subject", "duration[min]"])

        self.update_bt = QPushButton("update", self)
        self.days_spin_box = QSpinBox(self)
        self.days_spin_box.setMinimum(1)

        self.p_view_h_layout = QHBoxLayout()
        self.p_view_h_layout.addWidget(self.update_bt)
        self.p_view_h_layout.addWidget(self.days_spin_box)

        p_view_layout = QVBoxLayout(self.p_view)
        p_view_layout.setContentsMargins(0, 0, 0, 0)
        p_view_layout.addWidget(self.table)
        p_view_layout.addLayout(self.p_view_h_layout)
        self.stacked_widget.setCurrentWidget(self.p_view)

    def setup_menubar(self):
        self.action_theme = QAction("change theme")
        self.action_theme.setIcon(QIcon("./data/theme.ico"))
        self.menubar = self.menuBar()
        self.menu = self.menubar.addMenu("view")
        self.menu.addAction(self.action_theme)

    def setup_signals(self):
        self.days_spin_box.valueChanged.connect(self.check_updates)
        self.update_bt.clicked.connect(self.check_updates)
        self.action_theme.triggered.connect(self.toggle_theme)

    def toggle_theme(self):
        if self.theme == "dark":
            self.theme = "light"
        else:
            self.theme = "dark"

        with open(f"./data/{self.theme}.css", "r") as file:
            stylesheed = " ".join(file.readlines())
            self.setStyleSheet(stylesheed)

    def check_updates(self):
        self.model.clear()
        appointments = getCalendarEntries(self.days_spin_box.value())
        tupels = [x for x in zip(appointments['Start'], appointments['Subject'], appointments['Duration'])]

        for app in tupels:
            row = [QStandardItem(str(item)) for item in app]
            self.model.appendRow(row)


def getCalendarEntries(days=1):
    """
    Returns calender entries for days default is 1
    """
    Outlook = win32com.client.Dispatch("Outlook.Application")
    ns = Outlook.GetNamespace("MAPI")
    appointments = ns.GetDefaultFolder(9).Items
    appointments.Sort("[Start]")
    appointments.IncludeRecurrences = "True"
    today = datetime.datetime.today()
    begin = today.date().strftime("%d/%m/%Y")
    tomorrow = datetime.timedelta(days=days) + today
    end = tomorrow.date().strftime("%d/%m/%Y")
    appointments = appointments.Restrict("[Start] >= '" + begin + "' AND [END] <= '" + end + "'")
    events = {'Start': [], 'Subject': [], 'Duration': []}
    for a in appointments:
        events['Start'].append(datetime.datetime.strptime(str(a.Start), '%Y-%m-%d %H:%M:%S%z').time())
        events['Subject'].append(a.Subject)
        events['Duration'].append(a.Duration)
    return events


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = MainWindowUi()
    appointments = getCalendarEntries()
    sys.exit(app.exec_())

