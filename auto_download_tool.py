from PyQt4.QtCore import *
from PyQt4.QtGui import *
import sys
import os
import requests
import logging
import bs4
import win32process

logging.basicConfig(level=logging.INFO, format=' %(asctime)s - %(levelname)s - %(message)s')


class Ui_MainWindow(QDialog):
    def __init__(self, parent=None):
        super(Ui_MainWindow, self).__init__(parent)
        self.btnStartThunder = QPushButton(self)
        self.btnStartThunder.setGeometry(QRect(110, 390, 75, 23))
        self.btnStartThunder.setObjectName("btnStartThunder")
        self.btnStartThunder.setText('启动迅雷')

        self.btnCloseThunder = QPushButton(self)
        self.btnCloseThunder.setGeometry(QRect(10, 390, 75, 23))
        self.btnCloseThunder.setObjectName("btnCloseThunder")
        self.btnCloseThunder.setText('关闭迅雷')

        self.textBrowser = QTextBrowser(self)
        self.textBrowser.setGeometry(QRect(230, 40, 1000, 281))
        self.textBrowser.setObjectName("textBrowser")

        self.textBrowser_2 = QTextBrowser(self)
        self.textBrowser_2.setGeometry(QRect(10, 40, 201, 281))
        self.textBrowser_2.setObjectName("textBrowser_2")
        self.textBrowser_2.setText('使用提示：\n输入想抓取的网页地址，点击抓取按钮，会自动显示抓取到的所'
                                   '有的视频地址，再点击启动迅雷，然后选中视频地址，右键点击复制，或者'
                                   '按CTRL+C，迅雷会自动下载。在不需要迅雷的时候，可以点击‘关闭迅雷’按钮')

        self.edtWebAddress = QLineEdit(self)
        self.edtWebAddress.setGeometry(QRect(230, 360, 1000, 20))
        self.edtWebAddress.setObjectName("edtWebAddress")
        self.edtWebAddress.setText('请输入网页地址')

        self.edtThunderDir = QLineEdit(self)
        self.edtThunderDir.setGeometry(QRect(230, 390, 1000, 20))
        self.edtThunderDir.setObjectName("edtThunderDir")
        self.edtThunderDir.setText('请输入本机迅雷安装位置')

        self.btnCapture = QPushButton(self)
        self.btnCapture.setGeometry(QRect(110, 360, 75, 23))
        self.btnCapture.setObjectName("btnCapture")
        self.btnCapture.setText('抓取视频')

        self.setWindowTitle('电影视频机器手')
        self.setWindowFlags(Qt.Window)

        self.connect(self.btnCapture, SIGNAL("clicked()"), self.captureVideo)
        self.connect(self.btnStartThunder, SIGNAL("clicked()"), self.startThunder)
        self.connect(self.btnCloseThunder, SIGNAL('clicked()'), self.closeThunder)
        #创建迅雷进程句柄
        self.handle = []

    def captureVideo(self):
        logging.debug(self.edtWebAddress.text())
        try:
            resWeb = requests.get(self.edtWebAddress.text())
            logging.debug(resWeb.text)
        except:
            self.textBrowser.append('请输入网页地址')
            return
        # 判断请求是否正常返回
        try:
            resWeb.raise_for_status()
        except Exception as exc:
            self.textBrowser.append('网页请求未正常返回，请检查网页链接')
            return

        #切换编码格式，防止解码出现乱码
        if resWeb.encoding == 'ISO-8859-1':
            resWeb.encoding = 'gbk'

        bsWeb = bs4.BeautifulSoup(resWeb.text)
        bsTbody = bsWeb.select('tbody')
        logging.debug(resWeb.text)
        if len(bsTbody) > 0:
            for bsContext in bsTbody:
                logging.debug(bsContext)
                capturedText = bsContext.select('a')[0].get('href')
                logging.debug(capturedText)
                self.textBrowser.append(capturedText)
        else:
            self.textBrowser.setText('未抓取到视频信息，请检查网页链接')

    def startThunder(self):
        #如果在textBrowser里面已经选择了链接，则自动下载选中的文件
        thunderPath = os.path.join(self.edtThunderDir.text(), 'Thunder.exe')
        thunderParm = '-StartType:DesktopIcon'

        if os.path.exists(thunderPath): # 如果正确的设置迅雷安装位置
            logging.debug(thunderPath)
            self.handle = win32process.CreateProcess(thunderPath,
                                                thunderParm, None, None, 0, win32process.CREATE_NO_WINDOW,
                                                None, None, win32process.STARTUPINFO())
        else:
            self.textBrowser.append('迅雷设置不正确，使用默认迅雷位置：E:\Thunder Network\Thunder\Program\Thunder.exe')
            self.handle = win32process.CreateProcess("E:\Thunder Network\Thunder\Program\Thunder.exe",  '-StartType:DesktopIcon', None, None, 0, win32process.CREATE_NO_WINDOW, None, None, win32process.STARTUPINFO())

    def closeThunder(self):
        if len(self.handle) > 0:
            win32process.TerminateProcess(self.handle[0], 0)
        else:
            return None


app = QApplication(sys.argv)
form = Ui_MainWindow()
form.show()
app.exec_()
