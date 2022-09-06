import socket, sqlite3
from PyQt5 import QtCore

class TcpClient(QtCore.QObject):
    """
        отправка данных круга на компьютер табло,
        если получаем ответ, то обновляем базу и меняем цвет круга
    """
    sigSend = QtCore.pyqtSignal(str)
    def __init__(self, mat, parent= None):
        QtCore.QObject.__init__(self, parent)
        #self.s.settimeout(1)
        self.port = int(mat) * 1000
        self.addr = ""
        self.data = None

    def setData(self, addr, data):
        self.addr = addr
        self.data = data
        #print("addr = ", self.addr, "data = ", self.data)

    def run(self):
        print("send data")
        try:
            sock = socket.socket()
            sock.settimeout(2)
            sock.connect((self.addr, self.port))
            sock.send(bytes(self.data, "utf-8"))
            data = sock.recv(1024)
            sock.close()
            con = sqlite3.connect('baza.db')
            cur = con.cursor()
            try:
                sql = "UPDATE rounds SET sent = 1 WHERE id = " + data.decode()
                cur.execute(sql)
                con.commit()
                self.sigSend.emit(data.decode())
            except sqlite3.DatabaseError as err:
                print("error add to baze", err, data.decode())
            finally:
                cur.close()
                con.close()
            
            print("data tcp = ", data)
        except:
            print("err1")
            sock.close()

            
        
