import win32com.client as win32
from PyQt5 import QtGui, QtCore, QtWidgets
import sqlite3, os, sys, server, client
from ItemsQueue import *

excel = None
wb = None
currentWeight   = ""
currentCategory = ""
currentAge      = ""

def extractDataFromTuple(lst, tpl):
    if type(tpl) is tuple:
        for each in tpl:
            if type(each) is tuple:
                extractDataFromTuple(lst, each)
            else:
                if each != None:
                    try:
                        num = str(int(each))
                        lst.append(num)
                    except:
                        print("error data in table")
                    

class Communicate(QtCore.QObject):

    updateQueue = QtCore.pyqtSignal(str)

class SelectWindow(QtWidgets.QDialog):
    def __init__(self, parent=None):
        QtWidgets.QWidget.__init__(self, parent)
        global excel, wb, currentCategory, currentAge, currentWeight

        self.setWindowTitle(u"Выбор круга")
        self.setFixedSize(600, 400)
        self.startAddr = ""
        self.stopAddr = ""

        groupRButton = QtWidgets.QWidget()      #контейнер для радиокнопок
        
        scroll = QtWidgets.QScrollArea(self)
        scroll.setGeometry(0, 0, 300, 400)
        scroll.setVerticalScrollBarPolicy(0)
        scroll.setHorizontalScrollBarPolicy(1)

        nameWindow = QtWidgets.QWidget()      #окно предпросмотра фамилий, добавляемых в очередь
        nameWindow.setGeometry(0, 0, 300, 100)
        
        self.scrollName = QtWidgets.QScrollArea(self)
        self.scrollName.setGeometry(300, 0, 300, 300)
        self.scrollName.setVerticalScrollBarPolicy(0)
        self.scrollName.setHorizontalScrollBarPolicy(1)
        self.scrollName.setWidget(nameWindow)

        btnOk = QtWidgets.QPushButton(u"Отправить\nв очередь", self)
        btnOk.setGeometry(400, 325, 100, 50)
        btnOk.clicked.connect(self.sendToQueue)
    
        self.d = dict()
        print("c1")
        self.c = Communicate()
        print("c2")
        l = []
        try:
            for each in excel.Application.ActiveSheet.Names:
                print(each.Name)
                l.append(each.Name)
                if each.Name[4] == "К":
                    self.d[each.Name[4:]] = each
                elif each.Name[4:6] == "А_":
                    s = each.Name.split("_")
                    self.d["Подгруппа А, 1/" + s[1]] = each    
                elif each.Name[4:6] == "Б_":
                    s = each.Name.split("_")
                    self.d["Подгруппа Б, 1/" + s[1]] = each
                elif each.Name[4:6] == "АУ":
                    s = each.Name.split("_")
                    self.d["Утешительная подгруппа А, круг " + s[1]] = each
                elif each.Name[4:6] == "БУ":
                    s = each.Name.split("_")
                    self.d["Утешительная подгруппа Б, круг " + s[1]] = each
                elif each.Name[4] == "Ф":
                    self.d["Финал"] = each
        except:
            print("1")
            msg = QtWidgets.QMessageBox()
            msg.setText("Нет открытых файдов EXCEL")
            msg.move(0, 0)
            msg.exec_()

        lenghtDict = len(self.d)
        groupRButton.setGeometry(0, 0, 300, 30 * lenghtDict)
        count = 10
        for each in self.d.keys():
            rb = QtWidgets.QRadioButton(each, groupRButton)
            rb.move(20, count)
            rb.clicked.connect(self.selectEvent)
            count += 30
        
        scroll.setWidget(groupRButton)

        l=list(self.d.keys())
        sorted(l, key=lambda s: s[0])

        self.sportsmens = []
        self.currentRound = ""
        
        self.move(100, 100)
        self.show()
    
    def sendToQueue(self):
        conn = sqlite3.connect('baza.db')
        cursor = conn.cursor()
        print(self.sportsmens)
        if len(self.sportsmens) > 0:
            sql = "SELECT category, age, weight FROM sportsmens WHERE number = " +  str(self.sportsmens[0][2]) + " and name = '" \
                                                                                 +  self.sportsmens[0][0] + "' and region = '" \
                                                                                 +  self.sportsmens[0][1] + "'"
            category    = ""
            weight      = ""
            age         = ""
            c_a_v = None        #tuple with category, age, weight
            
            try:
                cursor.execute(sql)
                c_a_v = list(cursor.fetchall()[0])
            except sqlite3.DatabaseError as err:
                print("error:", err)

            try:
                cursor.execute("SELECT MAX(id) FROM rounds")
                rnd = cursor.fetchone()[0]
            except sqlite3.DatabaseError as err:
                print("error:", err)
            #текущий круг
            if rnd == None:
                rnd = 1
            else:
                rnd += 1
            try:
                cursor.execute("SELECT MAX(id) FROM queue")
                fight = cursor.fetchone()[0]
            except sqlite3.DatabaseError as err:
                print("error:", err)
            
            sql = "INSERT INTO rounds (category, age, weight, round) VALUES (?, ?, ?, ?)"

            c_a_v.append(self.currentRound)

            try:
                cursor.execute(sql, c_a_v)
            except sqlite3.DatabaseError as err:
                print("error:", err)
            conn.commit()

            for each in self.sportsmens:
                print("each = ", each)
                if len(each) == 5:
                    try:
                        sql = "SELECT docter FROM sportsmens WHERE id = " + str(each[4])
                        print(sql)
                        cursor.execute(sql)
                        val = cursor.fetchone()[0]
                        print("val = ", val)
                        each[3] = val
                        print("each = ", each)
                    except sqlite3.DatabaseError as err:
                        print("error:", err)
     
            number_of_fight = len(self.sportsmens)
            i = 0
            sql = "INSERT INTO queue (round, id_sportsmen_red, id_sportsmen_blue, note_red, note_blue) VALUES (?, ?, ?, ?, ?)"
            while i < number_of_fight:             
                if len(self.sportsmens[i]) > 0:
                    id_red = self.sportsmens[i][4]
                    if self.sportsmens[i][3] == 1:
                        note_red = "Сн.вр"
                    else:
                        note_red = ""
                else:
                    id_red = -1
                    note_red = ""
                if len(self.sportsmens[i + 1]) > 0:    
                    id_blue = self.sportsmens[i + 1][4]
                    if self.sportsmens[i + 1][3] == 1:
                        note_blue = "Сн.вр"
                    else:
                        note_blue = "" 
                else:
                    id_blue = -1
                    note_blue = ""
                try:
                    cursor.execute(sql, (rnd, id_red, id_blue, note_red, note_blue))
                except sqlite3.DatabaseError as err:
                    print("error:", err)
                i += 2
            conn.commit()       

        cursor.close()
        conn.close()
        self.c.updateQueue.emit(str(rnd))

    def note(self, state):
        conn = sqlite3.connect('baza.db')
        cursor = conn.cursor()
        
        sql =   "UPDATE sportsmens SET docter=" + str(int(state)) + " \
                 WHERE id = " + self.sender().objectName()
            
        try:
            cursor.execute(sql)
        except sqlite3.DatabaseError as err:
            print("error:", err)
        conn.commit()
        cursor.close()
        conn.close()
        
    def selectEvent(self, t):
        #global excel, wb, currentCategory, currentAge, currentWeight
        rb = self.findChildren(QtWidgets.QRadioButton)
        for each in rb:
            if each.isChecked():
                val = excel.Application.ActiveSheet.Range(self.d[each.text()]).Value
                self.currentRound = each.text()
                lst = []
                extractDataFromTuple(lst, val)
            
                conn = sqlite3.connect('baza.db')
                cursor = conn.cursor()
                sql =   """
                            SELECT name, region, number, docter, id FROM sportsmens 
                            WHERE number = ? and category = ? and weight = ? and age = ?
                        """
          
                self.sportsmens = []     #self.sportsmens = [name, region, number, docter, id]
                try:
                    for each in lst:
                        cursor.execute(sql, (each, currentCategory, currentWeight, currentAge))
                        v = cursor.fetchall()
                        
                        if len(v) == 1:
                            v = list(v[0])
                            #print("v = ", v, "\nlist(v) = ", list(v))
                        self.sportsmens.append(v)
                except sqlite3.DatabaseError as err:
                    print("error:", err)
                cursor.close()
                conn.close()
                print(self.sportsmens)
                count = 15
                wid = self.scrollName.takeWidget()
                lb = wid.findChildren(QtWidgets.QLabel)
                for each in lb:
                    each.setParent(None)
                    del each

                pair = 0
                lbl = QtWidgets.QLabel(u"Сн.вр", wid)
                lbl.setGeometry(243, 5, 30, 12)
                for each in self.sportsmens:
                    if len(each) > 0:
                        s = each[0] + ", " + each[1]        #Фамилия, имя + регион
                        chBox = QtWidgets.QCheckBox(wid)
                        chBox.setObjectName(str(each[4]))   #id спортсмена
                        chBox.clicked.connect(self.note)
                        chBox.move(250, count + 5)
                        if each[3] == 1:                    #Сн.вр
                            chBox.setChecked(True)
                    else:
                        s = u"св."
                    #print("s = ", s)
                    lbl = QtWidgets.QLabel(s, wid)
                    lbl.setGeometry(0,0,200,20)
                    lbl.move(20, count)
                    pair += 1
                    if pair == 2:
                        pair = 0
                        count += 20
                        lbl.setStyleSheet("QLabel{border-style: solid; border-bottom-width: 1px; border-bottom-color: black;}")
                    else:   
                        count += 20

                wid.setGeometry(0, 0, 300, count + 35)
                self.scrollName.setWidget(wid)    
                    
    def showEvent(self, e):
        global excel, wb
        

    def hideEvent(self, e):
        global excel, wb


class MainWindow(QtWidgets.QWidget):
    def __init__(self, parent=None):
        QtWidgets.QWidget.__init__(self, parent)

        global excel, wb

        self.currentY = 0
        #Получаем доступ к активному листу
        excel = win32.Dispatch("Excel.Application")
     
        try:
            count = excel.Workbooks.Count
        except:
            m = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Critical, "Ошибка","Возможно выделена ячейка таблицы двойным щелчком мыши. \
                                                                      Уберите выделение и запустите программу снова.")
            res = m.exec_()
            sys.exit()
   
        wb = []

        for i in range(1, count + 1):
            wb.append(excel.Workbooks.Item(i))
        
        self.cmbListFiles = QtWidgets.QComboBox(self)
        for each in wb:
            cort = each.Worksheets(u"взвешивание").Range("F7:F10").Value
            self.cmbListFiles.addItem(cort[0][0] + " " + cort[1][0] + " " + cort[2][0])

        self.cmbListFiles.currentIndexChanged.connect(self.chooseFile)

        self.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint)
        self.resize(600, 600)
        self.cmbListFiles.setGeometry(10, 30, 300, 20)

        #кнопка выбора спортсменов
        btnSelection = QtWidgets.QPushButton("Добавить\nв очередь", self)
        btnSelection.setGeometry(320, 20, 100, 40)
        btnSelection.clicked.connect(self.selectionOfAthletes)

        #область очереди
        self.queueWidndow = QtWidgets.QScrollArea(self)
        self.queueWidndow.setGeometry(10, 100, 580, 500)
        self.queueWidndow.setVerticalScrollBarPolicy(2)
        self.queueWidndow.setHorizontalScrollBarPolicy(1)

        #контейнер для виджетов в очереди
        self.queueContainer = QtWidgets.QWidget()

        #r = RoundFights()
        #l = QtWidgets.QLabel("hello", r)
        
        
        
        '''
        lblName = QtWidgets.QLabel("Фамилия, регион", self)
        lblName.setGeometry(10, 80, 230, 20)
        lblName.setAlignment(QtCore.Qt.AlignCenter)
        
        lblNote = QtWidgets.QLabel("Примечание", self)
        lblNote.setGeometry(240, 80, 80, 20)
        lblNote.setAlignment(QtCore.Qt.AlignCenter)
        
        lblSum = QtWidgets.QLabel("Сумма", self)
        lblSum.setGeometry(320, 80, 60, 20)
        lblSum.setAlignment(QtCore.Qt.AlignCenter)
        
        lblResult = QtWidgets.QLabel("Результат", self)
        lblResult.setGeometry(380, 80, 60, 20)
        lblResult.setAlignment(QtCore.Qt.AlignCenter)
        
        lblTime = QtWidgets.QLabel("Время", self)
        lblTime.setGeometry(440, 80, 60, 20)
        lblTime.setAlignment(QtCore.Qt.AlignCenter)
        
        lblSent = QtWidgets.QLabel("Отправлен", self)
        lblSent.setGeometry(500, 80, 60, 20)
        lblSent.setAlignment(QtCore.Qt.AlignCenter)
        '''
        
        #########################################################
        #                                                       #
        #                 создание базы данных                  #
        #                  с результатами боёв                  #
        #                                                       #
        #########################################################
        if not os.path.isfile('bazaFights.db'):
            conn = sqlite3.connect('bazaFights.db')
            cursor = conn.cursor()
            cursor.execute("""CREATE TABLE fights
                                (id INTEGER PRIMARY KEY AUTOINCREMENT,
                                fight INTEGER, result BLOB DEFAULT NONE)
                           """)
            conn.commit()
            cursor.close()
            conn.close()
        #########################################################
        #                                                       #
        #            создание основной базы данных              #
        #                                                       #
        #########################################################
        if not os.path.isfile('baza.db'):
            conn = sqlite3.connect('baza.db')
            cursor = conn.cursor()
            cursor.execute("""CREATE TABLE sportsmens
                            (id INTEGER PRIMARY KEY AUTOINCREMENT,
                            name TEXT, category TEXT, region TEXT,
                            age TEXT, weight TEXT, number INTEGER,
                            docter INTEGER DEFAULT 0,
                            UNIQUE (name, category, age, weight, number))
                            """)
            #########################################################################
            #                                                                       #
            #   //vinner                      -- победитель: 0 - не определён,      #
            #                                              1 - синий,               #
            #                                              2 - красный              #
            #   rate_red, rate_blue         -- баллы,                               #
            #   result_red, result_blue     -- счёт,                                #
            #   note_red, note_blue         -- примечания (Сн.вр., V, ЯП, и т.д.)   #
            #   sent                        -- 0 - не отправлено, 1 - отправлено,   #
            #   //accepted                    -- результат боя получен              #
            #                                                                       #
            #########################################################################
            cursor.execute("""CREATE TABLE queue
                                (id INTEGER PRIMARY KEY AUTOINCREMENT, round INTEGER,
                                id_sportsmen_red INTEGER DEFAULT -1, id_sportsmen_blue INTEGER DEFAULT -1,
                                note_red TEXT DEFAULT "", note_blue TEXT DEFAULT "")
                            """)
            cursor.execute("CREATE TABLE rounds (id INTEGER PRIMARY KEY AUTOINCREMENT, category TEXT, age TEXT, weight TEXT, round TEXT, sent INTEGER DEFAULT 0)")
            conn.commit()
            cursor.close()
            conn.close()
        ###########################################################
        #      открытие базы данных и заполнение значениями       #
        ###########################################################
        conn = sqlite3.connect('baza.db')
        cursor = conn.cursor()
        sql =   """
                    INSERT INTO sportsmens (name, region, category, age, weight, number)
                    VALUES (?, ?, ?, ?, ?, ?)
                """

        for each in wb:
            category =  each.Worksheets(u"взвешивание").Range("F7").Value
            age =       each.Worksheets(u"взвешивание").Range("F8").Value
            weight =    each.Worksheets(u"взвешивание").Range("F9").Value
            count = 0
            while True:
                count += 1
                rng = "B" + str(count)
                val = None
                try:
                    val = str(int(each.Worksheets(u"взвешивание").Range(rng).Value))
                    l = []
                    l.append(each.Worksheets(u"взвешивание").Range("D" + str(count)).Value)
                    l.append(each.Worksheets(u"взвешивание").Range("F" + str(count)).Value)
                    l.append(category)
                    l.append(age)
                    l.append(weight)
                    l.append(int(val))
                    #print(l)
                    try:
                        cursor.execute(sql, l)
                    except sqlite3.DatabaseError as err:
                        print("info:", err)
                    else:
                        conn.commit()      
                except:
                    pass
                if count > 100:
                    break

        try:
            cursor.execute("SELECT MAX(id) FROM rounds")
            rnd = cursor.fetchone()[0]
        except sqlite3.DatabaseError as err:
            print("error:", err)
        if rnd != None:
            for i in range(1, rnd + 1):
                self.updateQueue(str(i))
        cursor.close()
        conn.close()

        file = QtCore.QFile(u"ковер.txt")
        if not file.open(QtCore.QIODevice.ReadOnly):
            print("error open file")

        self.mat = file.read(1)
        file.close()
        #print(self.mat)
        
        #udp сервер для контроля связи и получения адреса
        self.th = QtCore.QThread()
        self.s = server.UdpServer(self.mat)
        
        self.s.moveToThread(self.th)
        self.th.started.connect(self.s.run)
        self.th.start()

        self.s.sigConn.connect(self.connection)
        self.lblConnect = QtWidgets.QLabel(u"Нет соединения", self)
        self.lblConnect.setGeometry(490, 20, 100, 40)
        self.lblConnect.setStyleSheet("QLabel{border-style: solid; border-width: 2px; border-color: red; color:red;border-radius:5px; font-size: 12px}")
        self.lblConnect.setAlignment(QtCore.Qt.AlignCenter)

        self.btnSendData = QtWidgets.QPushButton("", self)
        self.btnSendData.setGeometry(430, 20, 50, 40)
        self.btnSendData.clicked.connect(self.sendToTcp)
        self.btnSendData.setIcon(QtGui.QIcon("sending.png"))
        self.btnSendData.setIconSize(QtCore.QSize(30, 30))

        #tcp клиент для отправки данных на компьютер с табло
        self.tcpAddress = ""
        self.tcpClient = client.TcpClient(self.mat)
        self.tcpThread = QtCore.QThread()
        self.tcpClient.moveToThread(self.tcpThread)
        self.tcpThread.started.connect(self.tcpClient.run)
        self.tcpClient.sigSend.connect(self.sendEvent)
        #self.tcpThread.start()

        #tcp сервер для приема данных с компьютера с табло
        self.tcpServer = server.TcpServer(self.mat)
        self.tcpServerThread = QtCore.QThread()
        self.tcpServer.moveToThread(self.tcpServerThread)
        self.tcpServerThread.started.connect(self.tcpServer.run)
        self.tcpServer.sigData.connect(self.pixData)
        self.tcpServerThread.start()
        '''
        self.testLabel = QtWidgets.QLabel("test", self)
        self.testLabel.setGeometry(10, 140, 560, 80)
        self.testLabel.setFrameStyle(1)
        '''

    def pixData(self, d):
        print("pixdata len = ", len(d))
        pix = QtGui.QPixmap()
        print("1")
        #fight = str(d[0])
        #fight = chr(int(d[0]))    #.decode()
        
        fight = str(d[0])
        print("2", fight)
        data = d[1:]
        print("3")
        con = sqlite3.connect("bazaFights.db")
        cur = con.cursor()
        sql = "SELECT * FROM fights WHERE fight = " + fight
        try:
            cur.execute(sql)
            res = cur.fetchone()
            if res == None:
                cur.execute("INSERT INTO fights (fight, result) VALUES (?, ?)", (int(fight), data))
                con.commit()
            else:
                sql = "UPDATE fights SET result = ? WHERE fight = " + fight
                cur.execute(sql , (data,))
                con.commit()
        except sqlite3.DatabaseError as err:
            print("error = ", err)
        finally:
            cur.close()
            con.close()
        obj = self.findChild(Fight, fight)
        print("4")
        obj.setPix(data)
        print("5")
        
    def sendEvent(self, s):
        print("s = ", s)
        r = self.findChild(RoundFights, s)
        r.setSent()
        
    def connection(self, b, s):
        if b:
            #print("s = ", s)
            self.tcpAddress = s
            print("tcp")
            self.lblConnect.setText(u"Ковер " + str(self.mat)[2])
            self.lblConnect.setStyleSheet("QLabel{border-style: solid; border-width: 2px; border-color: green; color:green;border-radius:5px;font-size: 12px}")
            self.btnSendData.setEnabled(True)
        else:
            self.lblConnect.setText(u"Нет соединения")
            self.lblConnect.setStyleSheet("QLabel{border-style: solid; border-width: 2px; border-color: red; color:red;border-radius:5px;font-size: 12px}")
            self.btnSendData.setEnabled(False)

    def sendToTcp(self):
        if self.tcpAddress != "":
            conn = sqlite3.connect('baza.db')
            cursor = conn.cursor()
            sql =   "SELECT * FROM rounds WHERE sent = 0"                           
            try:
                cursor.execute(sql)
                data = cursor.fetchall()
            except:
                print("error ", sql)
                cursor.close()
                conn.close()
                return
            #print(data)
            cursor.close()
            conn.close()
            
            round_data = []
            
            
            for each in data:
                conn = sqlite3.connect('baza.db')
                cursor = conn.cursor()
                
                id_round = str(each[0])
                round_name = each[1] + ", " + each[2] + ", " + each[3] + ", " + each[4]
                #print(each)
                sql =   "SELECT id_sportsmen_red, id_sportsmen_blue, note_red, note_blue, id FROM queue WHERE round = " + str(each[0])
                #print(sql)
                
                try:
                    cursor.execute(sql)
                    data2 = cursor.fetchall()
                    print("data2 = ", data2)
                except sqlite3.Error as er:
                    print("error ", sql, er, er.args)
                    cursor.close()
                    conn.close()
                    return

                str_rounds = id_round + "=" + round_name + "="
                

                for i in data2:
                    sql_red =   "SELECT name, region, number FROM sportsmens WHERE id = " + str(i[0])
                    sql_blue =   "SELECT name, region, number FROM sportsmens WHERE id = " + str(i[1])
                    #print(data2)
                    try:
                        if i[0] != -1:
                            cursor.execute(sql_red)
                            d = cursor.fetchall()         
                            sportsmens_red = str(d[0][2]) + ":" + d[0][0] + "\n" + d[0][1]
                        else:
                            sportsmens_red = "св"
                        if i[1] != -1:
                            cursor.execute(sql_blue)
                            d = cursor.fetchall()
                            sportsmens_blue = str(d[0][2]) + ":" + d[0][0] + "\n" + d[0][1]
                        else:
                            sportsmens_blue = "св"
                        s = "<" + sportsmens_red + ";" + sportsmens_blue + ";" + i[2] + ";" + i[3] + ";" + str(i[4])
                        #print(s)
                    except sqlite3.Error as er:
                        print("error ", er, er.args)
                        cursor.close()
                        conn.close()
                        return
                    str_rounds = str_rounds + s
                print("str_rounds = ", str_rounds)

                cursor.close()
                conn.close()
                print("data = ", str_rounds)
                self.tcpClient.setData(self.tcpAddress, str_rounds)        
                self.tcpThread.start()
                self.tcpThread.quit()
                self.tcpThread.wait()

    #def mousePressEvent(self, e):
    #    print(e)
            
    def closeEvent(self, e):
        self.s.stopProcess()
        self.tcpServer.stopProcess()
        self.th.quit()
        self.tcpThread.quit()
        self.tcpServerThread.quit()
        self.th.wait()
        self.tcpThread.wait()
        self.tcpServerThread.wait()

    def updateQueue(self, rnd):
        conn = sqlite3.connect('baza.db')
        cursor = conn.cursor()

        #получение заголовков кругов
        try:
            cursor.execute("SELECT category, age, weight, round, sent FROM rounds WHERE id = " + rnd)
            val = cursor.fetchall()[0]
        except sqlite3.DatabaseError as err:
            print("error:", err)

        #print(val)
        r = RoundFights(val[0] + ", " + val[1] + ", "+ val[2] + "\n"+ val[3], self.queueContainer)
        r.setObjectName(rnd)
        if val[4] == 1:
            r.setSent()
        
        
        
        
        sql =   "SELECT * FROM queue WHERE round = " + rnd
        queue = None
        
        try:
            cursor.execute(sql)
            queue = cursor.fetchall()
        except sqlite3.DatabaseError as err:
            print("error:", err)
        for each in queue:
            data = []
            try:
                cursor.execute("SELECT name, region FROM sportsmens WHERE id = " + str(each[2]))
                val = cursor.fetchone()
                if val == None:
                    val = data.append("св")
                else:
                    data.append(val[0] + "\n" + val[1])
                cursor.execute("SELECT name, region FROM sportsmens WHERE id = " + str(each[3]))
                val = cursor.fetchone()
                if val == None:
                    val = data.append("св")
                else:
                    data.append(val[0] + "\n" + val[1])
                data += [each[4],each[5]]
            except sqlite3.DatabaseError as err:
                print("error:", err)
                
            r.addFight(data[0:4], str(each[0]))
            r.setGeometry(0, self.currentY, 580, r.currentHeight)
        self.currentY += r.currentHeight
        self.queueContainer.setGeometry(0, 0, 580, self.currentY)
        print(self.currentY)
        self.queueWidndow.setWidget(self.queueContainer)
        r.show()
        self.queueWidndow.ensureVisible(0, self.currentY)
        
        cursor.close()
        conn.close()

    def selectionOfAthletes(self):
        winSelect = SelectWindow()
        winSelect.setWindowModality(QtCore.Qt.ApplicationModal)
        winSelect.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint)
        winSelect.c.updateQueue.connect(self.updateQueue)
        res = winSelect.exec_()
        #winSelect.show()
        
    def chooseFile(self, index):
        global excel, wb, currentCategory, currentAge, currentWeight
        wb[index].Activate()
        wb[index].Worksheets(u"ход").Activate()

        currentCategory = wb[index].Worksheets(u"взвешивание").Range("F7").Value
        currentAge = wb[index].Worksheets(u"взвешивание").Range("F8").Value
        currentWeight = wb[index].Worksheets(u"взвешивание").Range("F9").Value
        print(currentCategory, currentAge, currentWeight)

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    window= MainWindow()
    window.show()
    sys.exit(app.exec_())
