from PyQt5.QtCore import Qt, QRect, QSize
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QCalendarWidget, QPushButton


class Ui_Calendar(object):
    def setupUi(self, calendar):
        font = QFont()
        font.setPointSize(10)
        calendar.setFixedSize(200, 200)
        calendar.setFont(font)
        calendar.setMouseTracking(True)
        calendar.setFocusPolicy(Qt.NoFocus)
        calendar.setGridVisible(True)
        calendar.setSelectionMode(QCalendarWidget.SingleSelection)
        calendar.setHorizontalHeaderFormat(QCalendarWidget.ShortDayNames)
        calendar.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
        calendar.setNavigationBarVisible(True)
        calendar.setDateEditAcceptDelay(0)

        calendar.setObjectName("calendarWidget")
        calendar.setStyleSheet(
            "/****************************************/\n"
            "QCalendarWidget QWidget {"
            "   alternate-background-color: rgba(0, 0, 0, 40);"
            "   outline:0;"
            "}"
            "QCalendarWidget QTableView QHeaderView::section {\n"
            "    /*font-family:Myriad*/\n"
            "}\n"
            "\n"
            "#qt_calendar_navigationbar {\n"
            "    /*background-color: rgba(250, 0, 0, 150);*/\n"
            "    background-color: rgb(240, 240, 240);\n"
            "    border: 1px solid rgb(173, 173, 173);\n"
            "}\n"
            "\n"
            "/****************************************/\n"
            "\n"
            "#qt_calendar_monthbutton,#qt_calendar_yearbutton {\n"
            "    font-size:14px;\n"
            "    margin-top: 5px;\n"
            "    margin-bottom: 5px;\n"
            "    color: rgb(0, 0, 0);\n"
            "    border:1;\n"
            # "   font-weight: normal;"
            "}\n"
            "#qt_calendar_monthbutton{\n"
            "    padding-left:auto;\n"
            "    padding-right:auto;\n"
            "    margin-left:auto;\n"
            "    margin-right:auto;\n"
            "}\n"
            "#qt_calendar_yearbutton {\n"
            "    padding-left:auto;\n"
            "    padding-right:auto;\n"
            "    margin-left:auto;\n"
            "    margin-right:auto;\n"
            "}\n"
            "#qt_calendar_monthbutton:hover,#qt_calendar_yearbutton:hover {\n"
            "    color: rgba(0, 0, 0,100);\n"
            "}\n"
            "\n"
            "#qt_calendar_yearedit {\n"
            "    font-size: 14px;\n"
            "    color: #000;\n"
            "    background: transparent;\n"
            "    margin-left:20;\n"
            "}\n"
            "#qt_calendar_yearedit::up-button { \n"
            "    image: url(Assets/up.svg);\n"
            "}\n"
            "\n"
            "#qt_calendar_yearedit::down-button { \n"
            "    image: url(Assets/down.svg);\n"
            "}\n"
            "\n"
            "#qt_calendar_yearedit::down-button, \n"
            "#qt_calendar_yearedit::up-button {\n"
            "    width:20;\n"
            "    border:1px;\n"
            "\n"
            "}\n"
            "#qt_calendar_yearedit::down-button:hover, \n"
            "#qt_calendar_yearedit::up-button:hover {\n"
            "    background-color: rgba(0, 0, 0,40);\n"
            "}\n"
            "#qt_calendar_monthbutton::menu-indicator {\n"
            "   image:none"
            "}\n"
            "\n"
            "/****************************************/\n"
            "\n"
            "#qt_calendar_prevmonth {\n"
            "    padding-left:15px;\n"
            "    padding-right:8;\n"
            "    image: url(Assets/back.svg);\n"
            "}\n"
            "#qt_calendar_nextmonth {\n"
            "   margin-left:auto;\n"
            "   padding-left:13;\n"
            "   padding-right:10;\n"
            "   image: url(Assets/next.svg);\n"
            "}\n"
            "\n"
            "#qt_calendar_prevmonth,#qt_calendar_nextmonth {\n"
            "    border:1;\n"
            "    qproperty-icon:none;\n"
            "}\n"
            "#qt_calendar_prevmonth:hover, \n"
            "#qt_calendar_nextmonth:hover {\n"
            "    background-color: rgba(0, 0, 0, 40);\n"
            "}\n"
            "\n"
            "/****************************************/\n"
            "#qt_calendar_calendarview {\n"
            "   border: 1px solid rgb(173, 173, 173);\n"
            "   border-top: 0px;\n"
            "}\n"
            "QCalendarWidget QTableView::item:selected {"
            "    background-color: #ff0000; /* Красный фон */"
            "    color: #ffffff; /* Белый текст */"
            "}"
            "QCalendarWidget QAbstractItemView QWidget:enabled:hover QPushButton:!disabled {"
            "    background-color: #337ab7;"
            "    color: #ffffff;"
            "}"
        )

        self.pushButton = QPushButton(calendar)
        self.pushButton.setGeometry(QRect(200, 0, 40, 30))
        self.pushButton.setObjectName("pushButton")
        self.pushButton.setStyleSheet(
            "\n"
            "QPushButton {\n"
            "    background:transparent;\n"
            "    image: url(Assets/today.svg);\n"
            "}\n"
            "QPushButton:hover {\n"
            "    \n"
            "    background-color: rgba(0, 0, 0, 20);\n"
            "}"
        )
