from PyQt5 import QtNetwork as n
from PyQt5 import QtCore, QtWidgets
import socket

class UdpServer(QtCore.QObject):
    #def __init__(self, parent=None):
    sigConn = QtCore.pyqtSignal(bool, str)
    def __init__(self, mat, parent= None):
        QtCore.QObject.__init__(self, parent)
        self.s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        self.s.settimeout(3)
        self.s.bind(("", int(mat) * 2000 + 1))
        self.process = True
        self.mat = mat
        #print("self.mat = ", self.mat)

    def run(self):
        while self.process:
            try:
                b = self.s.recvfrom(256)
                print(b)
                if b[0] == self.mat:
                    addr = (b[1][0], int(self.mat) * 2000)
                    #print("addr = ", addr)
                    self.s.sendto(b[0], addr)
                    #print("sended")
                    self.sigConn.emit(True, b[1][0])
                    #strk = "signal = " + b[1][0] + "\n"
                    #print(strk)
            except socket.timeout:
                self.sigConn.emit(False, "")
            except ConnectionResetError:
                #print("err")
                pass
                
        self.s.close()
        
    def stopProcess(self):
        self.process = False

class TcpServer(QtCore.QObject):
    sigData = QtCore.pyqtSignal(bytes)
    def __init__(self, mat, parent= None):
        QtCore.QObject.__init__(self, parent)
        self.s = socket.socket()
        self.s.settimeout(3)
        self.s.bind(("", int(mat) * 2000 + 3))
        self.s.listen(1)
        self.process = True

    def run(self):
        while self.process:
            try:
                conn, addr = self.s.accept()
                print("tcp accept", conn, addr)
                conn.settimeout(0.3)
                try:
                    d = b''
                    while True:
                        data = conn.recv(1024)
                        #print("data = ", data)
                        if not data: break
                        d += data
                    
                    #print(len(d))
                    self.sigData.emit(d)
                except socket.timeout:
                    print("tcp conn timeout")
                    #pass
                    
                finally:
                    conn.close()
                
            except socket.timeout:
                print("tcp server timeout")
                pass

        self.s.close()
        
    def stopProcess(self):
        self.process = False
        
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    UServer = UdpServer("1") 

    UServer.run()
    #window.show()
    sys.exit(app.exec_())
