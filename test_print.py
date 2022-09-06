import sys
from PyQt5.QtWidgets import QApplication,QMainWindow,QLabel,QSizePolicy,QAction
from PyQt5.QtPrintSupport import QPrintDialog,QPrinter
from PyQt5.QtGui import QImage,QIcon,QPixmap

class MainWindow(QMainWindow):
    def  __init__(self,parent=None):
        super(MainWindow, self).__init__(parent)

if __name__ == '__main__':
    app=QApplication(sys.argv)
    main=MainWindow()
    main.show()
    sys.exit(app.exec_())
