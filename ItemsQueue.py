from PyQt5 import QtGui, QtCore, QtWidgets
from PyQt5.QtPrintSupport import QPrintDialog,QPrinter
import sqlite3

class ViewFight(QtWidgets.QSplashScreen):
    def __init__(self, pix, name, parent=None):
        QtWidgets.QSplashScreen.__init__(self, parent)
        self.setGeometry(50, 50, pix.width(), pix.height())
        self.setPixmap(pix)
        self.pix = pix
        self.objName = name

    def mousePressEvent(self, e):
        if e.button() == 1:
            self.destroy()
        else:
            # Создание объекта печати изображения
            printer = QPrinter()
            printer.setOrientation(QPrinter.Landscape)
            # Откроется окно печати
            printDialog = QPrintDialog(printer, self)
            if printDialog.exec_():
                conn = sqlite3.connect('baza.db')
                cursor = conn.cursor()
                try:
                    cursor.execute("SELECT round FROM queue WHERE id = " + self.objName)
                    rnd = cursor.fetchone()[0]
                    cursor.execute("SELECT * FROM rounds WHERE id = " + str(rnd))
                    category = cursor.fetchone()
                except sqlite3.DatabaseError as err:
                    print("error:", err)
                finally:
                    cursor.close()
                    conn.close()
                    
                painter = QtGui.QPainter()
                if painter.begin(printer):
                    font = QtGui.QFont()
                    font.setPixelSize(20)
                    painter.setFont(font)
                    # Реализованное окно просмотра
                    rect = painter.viewport()
                    #print(rect)
                    # Получите размер картинки
                    size = self.pix.size()
                    #print(size)
                    size.scale(rect.size(), QtCore.Qt.KeepAspectRatio)
                    #print(size)
                    # Установить свойства окна просмотра
                    painter.setViewport(rect.x() + 20, rect.y(), size.width() - 20, size.height())
                    #print(painter.viewport())

                    # Установите размер окна под размер картинки и нарисуйте картинку в окне
                    painter.setWindow(self.pix.rect())
                    #print("1")
                    text = "ПРОТОКОЛ\n боя по рукопашному бою\n" + category[1] + ", " \
                                                                 + category[2] + ", " \
                                                                 + category[3] + ", " \
                                                                 + category[4]
                    #print(text)
                    painter.drawText(QtCore.QRectF(20, 0, 1073, 100), text,
                                     QtGui.QTextOption(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter))
                    painter.drawPixmap(0, 100, self.pix)
                    
                    painter.end()
                    #painter.drawText(0, 0, "hello")
                         
class Fight(QtWidgets.QWidget):
    def __init__(self, l, parent=None):
        QtWidgets.QWidget.__init__(self, parent)
        self.lblRed = QtWidgets.QLabel(l[0], self)
        self.lblRed.setGeometry(2, 0, 200, 40)
        self.lblRed.setStyleSheet("QLabel{border-style: solid; border-right-width: 1px; border-bottom-color: black;}")
        
        self.lblBlue = QtWidgets.QLabel(l[1], self)
        self.lblBlue.setGeometry(2, 40, 200, 40)
        self.lblBlue.setStyleSheet("QLabel{border-style: solid; border-right-width: 1px; border-bottom-color: black;}")
        
        self.lblNoteRed = QtWidgets.QLabel(l[2], self)
        self.lblNoteRed.setGeometry(202, 0, 60, 40)
        self.lblNoteRed.setStyleSheet("QLabel{border-style: solid; border-right-width: 1px; border-bottom-color: black;}")
        self.lblNoteRed.setAlignment(QtCore.Qt.AlignCenter)
        
        self.lblNoteBlue = QtWidgets.QLabel(l[3], self)
        self.lblNoteBlue.setGeometry(202, 40, 60, 40)
        self.lblNoteBlue.setStyleSheet("QLabel{border-style: solid; border-right-width: 1px; border-bottom-color: black;}")
        self.lblNoteBlue.setAlignment(QtCore.Qt.AlignCenter)

        self.lblFight = QtWidgets.QLabel("hellow", self)
        self.lblFight.setGeometry(0, 0, 560, 80)
        self.lblFight.setVisible(False)
        
    def paintEvent(self, e):
        painter = QtGui.QPainter(self)
        painter.drawLine(0, 40, 560, 40)
        pen = QtGui.QPen()
        pen.setWidth(3)
        painter.setPen(pen)
        painter.drawLine(0, 79, 560, 79)

    def mousePressEvent(self, e):
        if self.lblFight.isVisible():
            con = sqlite3.connect("bazaFights.db")
            cur = con.cursor()
            sql = "SELECT * FROM fights WHERE fight = " + self.objectName()
            try:
                cur.execute(sql)
                res = cur.fetchone()[2]
                pix = QtGui.QPixmap()
                pix.loadFromData(res, "PNG")
                view = ViewFight(pix, self.objectName(), self)
                view.show()
            except sqlite3.DatabaseError as err:
                print("error = ", err)
            finally:
                cur.close()
                con.close()

    def setPix(self, data):
        pix = QtGui.QPixmap()
        pix.loadFromData(data, "PNG")
        self.lblFight.setPixmap(pix.scaled(self.lblFight.width(), self.lblFight.height()))
        self.lblFight.setVisible(True)
  
class RoundFights(QtWidgets.QWidget):
    def __init__(self, rnd, parent=None):
        QtWidgets.QWidget.__init__(self, parent)
        self.rnd = rnd
        self.blue = ""
        self.note_red = ""
        self.note_blue = ""
        self.currY = 40
        self.currentHeight = 60
        self.sent = False   #флаг отправки данных на табло

    def setSent(self):
        self.sent = True
        self.update()

    def addFight(self, l, fight):
        f = Fight(l, self)
        f.setGeometry(0, self.currY, 560, 80)
        f.setObjectName(fight)
        self.currY += 82
        self.currentHeight += 82
        con = sqlite3.connect("bazaFights.db")
        cur = con.cursor()
        sql = "SELECT * FROM fights WHERE fight = " + fight
        try:
            cur.execute(sql)
            res = cur.fetchone()
            if res != None:
                f.setPix(res[2])
        except sqlite3.DatabaseError as err:
            print("error = ", err)
        finally:
            cur.close()
            con.close()
        
    def paintEvent(self, e):
        painter = QtGui.QPainter(self)
        if self.sent:
            painter.setBrush(QtCore.Qt.white)
        painter.drawRoundedRect(0, 0, self.width() - 20, self.height() - 1, 15.0, 15.0)
        #painter.setPen(QtCore.Qt.black)
        painter.setFont(QtGui.QFont("Arial", 10))
        painter.drawText(0, 0, self.width(), 40, QtCore.Qt.AlignCenter, self.rnd)
        
        
        painter.drawLine(0, 40, self.width() - 20, 40)
        
                
