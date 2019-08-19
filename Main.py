import sys

from PySide2.QtCore import QUrl
from PySide2.QtCore import Qt
from PySide2.QtWidgets import QApplication
from PySide2.QtWidgets import QMainWindow
from PySide2.QtWidgets import QSplitter

from Util import MySurpac
from Util import MyTabWidget
from Util import MyTreeWidget
from Util import MyWebView

if __name__ == "__main__":
    # 启动应用
    app = QApplication(sys.argv)

    # 生成窗口并配置
    mainWindow = QMainWindow()
    mainWindow.resize(640, 480)
    mainWindow.setWindowTitle('中矿智信三维管控平台客户端')

    # 生成并配置tab组件
    tabWdgt = MyTabWidget()

    # 嵌入浏览器界面
    webView = MyWebView(tabWdgt)
    webView.load(QUrl("http://www.ifeng.com"))
    tabWdgt.createTabItem(webView, '首页')

    # 嵌入surpac界面
    pname = 'surpac2'
    cmd = "C:/Program Files (x86)/GEOVIA/Surpac/69/nt_i386/bin/surpac2.exe"
    spTitle = 'GEOVIA Surpac'
    pids = MySurpac.getPidsFromPName(pname)
    MySurpac.killProcess(pids)
    pid = MySurpac.startProcess(cmd)
    hwnd = MySurpac.getTheMainWindow(pid, spTitle)
    ports = MySurpac.getPortsFromPid(pid)
    surpacWdgt = MySurpac.convertWndToWidget(hwnd)

    # 嵌入Tree界面
    treeWdgt = MyTreeWidget(ports[0])

    # 分割窗口
    splitter = QSplitter()
    splitter.setOrientation(Qt.Horizontal)
    splitter.addWidget(surpacWdgt)
    splitter.addWidget(treeWdgt)
    tabWdgt.createTabItem(splitter, '三维设计')

    mainWindow.setCentralWidget(tabWdgt)
    mainWindow.showMaximized()

    # 退出应用
    sys.exit(app.exec_())
