import sys
import traceback

from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QLabel, QVBoxLayout, QWidget, \
    QMessageBox, QProgressDialog, QTextEdit, QStatusBar, QHBoxLayout, QLineEdit
from PyQt5.QtCore import Qt, QDate
from PyQt5 import QtWidgets, QtCore
from Storage_Final_NEW import *


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.子部件_文件選擇 = 子部件_文件選擇()
        self.日期選擇 = DateRangePicker()
        self.UI()

    def UI(self):
        # 介面標題
        # 設定視窗位置與大小
        self.setWindowTitle("銷貨與庫存核對")
        self.setGeometry(400, 300, 600, 200)

        # 建立 QLabel 用於顯示背景圖片
        background_label = QLabel(self)

        # 加載背景圖片
        pixmap = QPixmap(':/Background.jpg')

        # 設置 QLabel 的尺寸和背景圖片
        # self.width() 和 self.height() 分別返回視窗的寬度和高度，這樣可以確保背景圖片的大小與視窗相符
        # (x, y) 設置為 (0, 0)，它將位於視窗的左上角
        background_label.setGeometry(0, 0, self.width(), self.height())
        background_label.setPixmap(pixmap)
        background_label.setScaledContents(True)

        # 將背景圖片置於最底層
        background_label.lower()

        Mainlayout = QVBoxLayout()

        self.執行鈕 = QPushButton('執行')
        self.執行鈕.setStyleSheet("font-size: 16px;font-family: 新細明體;font-weight: bold")
        self.執行鈕.setFixedSize(150, 30)
        self.執行鈕.clicked.connect(self.Run)

        title_label = QLabel("銷貨與庫存核對")
        title_label.setStyleSheet("font-size: 24px; font-weight: bold;")

        Mainlayout.addWidget(title_label, alignment=Qt.AlignCenter | Qt.AlignTop)
        Mainlayout.addWidget(self.日期選擇.核對日期框)
        Mainlayout.addLayout(self.子部件_文件選擇.文件選擇_422)
        Mainlayout.addLayout(self.子部件_文件選擇.文件選擇_GDS)
        Mainlayout.addWidget(self.執行鈕)
        # Mainlayout.setAlignment(Qt.AlignCenter | Qt.AlignTop)
        # 創建一個 QWidget 作為佈局的容器
        container = QWidget(self)
        container.setLayout(Mainlayout)
        self.setCentralWidget(container)

    def Run(self):
        try:
            if not self.子部件_文件選擇.檔案選擇確認_422 or not self.子部件_文件選擇.檔案選擇確認_GDS:
                QMessageBox.warning(self, "警告", "尚有檔案未選取！")
                return
            月結客戶 = ['2E-營邦', '7O-兆赫', '44-Asentria', '5Q-B&B', '0K-KFA']
            輸出用檔案 = 檔案核對(self.子部件_文件選擇.檔案選擇確認_GDS,
                                self.子部件_文件選擇.檔案選擇確認_422,
                                月結客戶,
                                self.日期選擇.核對日期)
            Excel格式輸出(輸出用檔案)
            QMessageBox.information(self, '結果', '文件處理完成!')
        except Exception as e:
            traceback.print_exc()
            QMessageBox.warning(self, "文件讀取與輸出錯誤", f"發生錯誤：{e}\n請確認檔案選取是否正確!")




class 子部件_文件選擇(QWidget):
    def __init__(self):
        super().__init__()
        self.檔案選擇確認_422 = False
        self.檔案選擇確認_GDS = False
        self.UI()

    def UI(self):
        self.setStyleSheet("""
                                    QLabel {
                                        color: red;
                                        font-size: 20px;
                                    }
                                """)

        self.文件選擇_422 = QHBoxLayout()
        self.文件選擇_GDS = QHBoxLayout()

        self.提示字串_422 = QLabel('庫存報表(422) :')
        self.提示字串_GDS = QLabel('銷貨報表(GDS) :')
        self.提示字串_422.setStyleSheet("font-size: 16px;font-family: 新細明體;font-weight: bold")
        self.提示字串_GDS.setStyleSheet("font-size: 16px;font-family: 新細明體;font-weight: bold")

        self.文件選擇鈕_422 = self.button_style()
        self.文件選擇鈕_GDS = self.button_style()
        self.文件選擇鈕_422.clicked.connect(self.selectFile_422)
        self.文件選擇鈕_GDS.clicked.connect(self.selectFile_GDS)

        self.文件顯示框_422 = self.lineEdit_style()
        self.文件顯示框_GDS = self.lineEdit_style()

        self.文件選擇_422.addWidget(self.提示字串_422)
        self.文件選擇_422.addWidget(self.文件顯示框_422)
        self.文件選擇_422.addWidget(self.文件選擇鈕_422)
        self.文件選擇_GDS.addWidget(self.提示字串_GDS)
        self.文件選擇_GDS.addWidget(self.文件顯示框_GDS)
        self.文件選擇_GDS.addWidget(self.文件選擇鈕_GDS)

    def lineEdit_style(self):
        通用lineEdit設定 = QLineEdit()
        通用lineEdit設定.setStyleSheet("font-size: 16px;font-family: 新細明體;font-weight: bold")
        通用lineEdit設定.setFixedSize(350, 30)
        通用lineEdit設定.setPlaceholderText('請選擇或貼上文件路徑')
        return 通用lineEdit設定

    def button_style(self):
        通用按鈕設定 = QPushButton('...')
        通用按鈕設定.setStyleSheet("font-size: 16px;font-family: 新細明體;font-weight: bold")
        通用按鈕設定.setFixedSize(30, 30)
        return 通用按鈕設定

    def selectFile_422(self):
        文件選擇視窗 = QFileDialog()

        # 使用變數 file_path 來接收文件路徑，而 _ 變數表示我們不關心文件類型
        文件路徑, _ = 文件選擇視窗.getOpenFileName(self, "選擇檔案")
        self.檔案選擇確認_422 = 文件路徑
        self.文件顯示框_422.setText(f"{文件路徑}")

    def selectFile_GDS(self):
        文件選擇視窗 = QFileDialog()

        # 使用變數 file_path 來接收文件路徑，而 _ 變數表示我們不關心文件類型
        文件路徑, _ = 文件選擇視窗.getOpenFileName(self, "選擇檔案")
        self.檔案選擇確認_GDS = 文件路徑
        self.文件顯示框_GDS.setText(f"{文件路徑}")


class DateRangePicker(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        初始日期 = QtCore.QDate.currentDate()

        # 這個是供函式判斷用的日期，設定為字串及特定格式
        self.核對日期 = 初始日期.toString("yyyy/MM/dd")

        # 創建起始日期的 QDateEdit
        # QtWidgets.QDateEdit() 函數內可設置預設日期
        # 這個是介面上的日期選擇框，讓使用者看到和修改時間，變動時同步修改上面3個變數
        self.核對日期框 = self.General_createDateEdit(初始日期)

        # 連接按鈕的點擊事件到槽函數
        self.核對日期框.dateChanged.connect(self.Changdate)

    def General_createDateEdit(self, date):
        date_edit = QtWidgets.QDateEdit(date)
        date_edit.setCalendarPopup(True)
        date_edit.setFixedWidth(115)
        date_edit.setStyleSheet("font-size: 18px;font-Family: Times New Roman")
        return date_edit

    def Changdate(self):
        self.核對日期 = self.核對日期框.date().toString("yyyy/MM/dd")
        print(self.核對日期)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
