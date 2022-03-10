from PyQt5 import QtCore, QtWidgets, QtWebEngineWidgets
import sys

port = None

class Renderer(object):
    def __init__(self, MainWindow):
        super().__init__()
        global port
        self.MainWindow = MainWindow
        self.MainWindow.setObjectName("CorelDraw Macro Tools")
        self.MainWindow.setWindowTitle("CorelDraw Macro Tools")
        self.MainWindow.resize(1000, 600)
        self.MainWindow.setMinimumSize(800, 400)
        self.port = port
        print(self.port)

        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
     
        self.init_webView()

        
    def init_webView(self):
        
        self.webView = QtWebEngineWidgets.QWebEngineView(self.centralwidget)
        self.webView.setGeometry(QtCore.QRect(0, 0, self.MainWindow.frameGeometry().width(), self.MainWindow.frameGeometry().height()))
        self.webView.setUrl(QtCore.QUrl(f"http://localhost:{self.port}"))
        self.webView.page().profile().clearHttpCache()
        self.webView.setObjectName("webView")
        self.MainWindow.setCentralWidget(self.centralwidget)
        
    def resizeWebView(self):
        self.webView.setGeometry(QtCore.QRect(0, 0, self.MainWindow.frameGeometry().width(), self.MainWindow.frameGeometry().height()))



class Window(QtWidgets.QMainWindow):
    resized = QtCore.pyqtSignal()

    
    def  __init__(self):
        super().__init__()
        self.renderer = Renderer(MainWindow=self)

        
        self.resized.connect(lambda: self.renderer.resizeWebView())
        

    def resizeEvent(self, event):
        self.resized.emit()
        return None

def makeWindow(given_port):
    
    global port
    port = given_port
    
    print(given_port)
    
    app = QtWidgets.QApplication(sys.argv)

    MainWindow = Window()

    MainWindow.show()

    sys.exit(app.exec_())

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    
    MainWindow = Window(5322)
    
    MainWindow.show()
    
    sys.exit(app.exec_())
