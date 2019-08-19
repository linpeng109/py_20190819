# encoding:utf-8
import datetime
import os
import signal
import socket
import subprocess
import time

import win32con
import win32gui
import win32process

from PySide2.QtCore import QThread, Qt
from PySide2.QtGui import QWindow
from PySide2.QtWebEngineWidgets import QWebEngineSettings
from PySide2.QtWebEngineWidgets import QWebEngineView
from PySide2.QtWidgets import QHBoxLayout
from PySide2.QtWidgets import QTabWidget
from PySide2.QtWidgets import QTreeWidget
from PySide2.QtWidgets import QTreeWidgetItem
from PySide2.QtWidgets import QWidget


class MyThread(QThread):
    # port: int

    def __init__(self, port, item):
        super().__init__()
        self.port = port
        self.item = item

    def run(self):
        print(self.item.text(0))
        print(self.port)
        socket = MySocketClient(int(self.port), 'gbk')
        extend = "EXIT GRAPHICS"
        message = 'RCTL\n' + 'TCLSCRIPTBEGIN\n' \
                  + 'set status [ SclFunction ' \
                  + extend + ' {} ]' \
                  + 'TCLSCRIPTEND\n'
        result = socket.sendMsg(message)
        print(result.decode(socket.ENCODE))
        socket.closeSocket()
        pass


class MyTreeWidget(QTreeWidget):
    def __init__(self, port):
        super(MyTreeWidget, self).__init__()
        self.root = QTreeWidgetItem(self)
        self.port = port
        self.setColumnCount(1)
        self.setHeaderHidden(True)
        self.itemClicked.connect(self.onItemClicked)

        root = QTreeWidgetItem(self)
        root.setText(0, '中矿智信SURPAC扩展命令（V2）')
        child1 = QTreeWidgetItem(root)
        child1.setText(0, '露采地质')

        child11 = QTreeWidgetItem(child1)
        child11.setText(0, '创建月度炮孔数据库')
        child11.setText(1, 'aaaa')
        child11.setText(2, 'bbbb')
        child11.setText(3, 'cccc')
        child12 = QTreeWidgetItem(child1)
        child12.setText(0, '导入炮孔数据')
        child13 = QTreeWidgetItem(child1)
        child13.setText(0, '圈定爆堆边界')
        child14 = QTreeWidgetItem(child1)
        child14.setText(0, '圈定矿岩界线')
        child15 = QTreeWidgetItem(child1)
        child15.setText(0, '矿岩多边形算量')
        child16 = QTreeWidgetItem(child1)
        child16.setText(0, '炮孔信息处理')
        child17 = QTreeWidgetItem(child1)
        child17.setText(0, '矿岩多边形信息写入爆堆数据库')
        child18 = QTreeWidgetItem(child1)
        child18.setText(0, '当前目录矿岩多边形信息写入爆堆数据库')
        child19 = QTreeWidgetItem(child1)
        child19.setText(0, '核验爆堆数据库')
        child1a = QTreeWidgetItem(child1)
        child1a.setText(0, '数据库爆堆信息写入爆堆文件')
        child1b = QTreeWidgetItem(child1)
        child1b.setText(0, '报表统计')
        child1c = QTreeWidgetItem(child1)
        child1c.setText(0, '爆堆估值赋值')
        child1d = QTreeWidgetItem(child1)
        child1d.setText(0, '探采对比')
        child1e = QTreeWidgetItem(child1)
        child1e.setText(0, '块模型处理')
        child1f = QTreeWidgetItem(child1)
        child1f.setText(0, '在工作目录下搜索文件合并')
        child1g = QTreeWidgetItem(child1)
        child1g.setText(0, '绘制台阶潜孔取样及矿体圈定图')
        child1h = QTreeWidgetItem(child1)
        child1h.setText(0, '创建目录')
        child2 = QTreeWidgetItem(root)
        child2.setText(0, '露采测量')
        child21 = QTreeWidgetItem(child2)
        child21.setText(0, '导入现状收测数据')
        child22 = QTreeWidgetItem(child2)
        child22.setText(0, '调入已有收测范围线')
        child23 = QTreeWidgetItem(child2)
        child23.setText(0, '设置收测范围线线串好按分层号')
        child24 = QTreeWidgetItem(child2)
        child24.setText(0, '施工范围线赋属性')
        child25 = QTreeWidgetItem(child2)
        child25.setText(0, '计算填挖方量并形成汇总表')
        child26 = QTreeWidgetItem(child2)
        child26.setText(0, '图层内选择线串范围创建三维')
        child27 = QTreeWidgetItem(child2)
        child27.setText(0, '两个段之间创建三维示坡线')

        child3 = QTreeWidgetItem(root)
        child3.setText(0, '露采采矿')
        child31 = QTreeWidgetItem(child3)
        child31.setText(0, '初始化')
        child32 = QTreeWidgetItem(child3)
        child32.setText(0, '调入已有数据')
        child33 = QTreeWidgetItem(child3)
        child33.setText(0, '台阶推荐条带')
        child34 = QTreeWidgetItem(child3)
        child34.setText(0, '台阶开拓条带')
        child35 = QTreeWidgetItem(child3)
        child35.setText(0, '为采掘条带赋编号和单位')
        child36 = QTreeWidgetItem(child3)
        child36.setText(0, '删除采掘条带')
        child37 = QTreeWidgetItem(child3)
        child37.setText(0, '采掘条带算量')
        child38 = QTreeWidgetItem(child3)
        child38.setText(0, '周计划')

        child4 = QTreeWidgetItem(root)
        child4.setText(0, '地采地质')
        child41 = QTreeWidgetItem(child4)
        child41.setText(0, '数据处理')
        child42 = QTreeWidgetItem(child4)
        child42.setText(0, '品位控制模型更新')
        child43 = QTreeWidgetItem(child4)
        child43.setText(0, '回采模型更新')
        child44 = QTreeWidgetItem(child4)
        child44.setText(0, '绘图')

        child5 = QTreeWidgetItem(root)
        child5.setText(0, '地采测量')
        child51 = QTreeWidgetItem(child5)
        child51.setText(0, '数据处理')
        child52 = QTreeWidgetItem(child5)
        child52.setText(0, '品位控制模型更新')
        child53 = QTreeWidgetItem(child5)
        child53.setText(0, '回采模型更新')
        child54 = QTreeWidgetItem(child5)
        child54.setText(0, '绘图')

        child6 = QTreeWidgetItem(root)
        child6.setText(0, '地采采矿')
        child61 = QTreeWidgetItem(child6)
        child61.setText(0, '数据准备')
        child62 = QTreeWidgetItem(child6)
        child62.setText(0, '爆破准备')
        child63 = QTreeWidgetItem(child6)
        child63.setText(0, '爆破参考边界')
        child64 = QTreeWidgetItem(child6)
        child64.setText(0, '单排巷（单井或VCR）')
        child65 = QTreeWidgetItem(child6)
        child65.setText(0, '单排巷（双井或VCR）')
        child66 = QTreeWidgetItem(child6)
        child66.setText(0, '双排巷（单井或VCR）')
        child67 = QTreeWidgetItem(child6)
        child67.setText(0, '双排巷（双井或VCR）')
        child68 = QTreeWidgetItem(child6)
        child68.setText(0, '井孔编辑')
        child69 = QTreeWidgetItem(child6)
        child69.setText(0, '单体')
        child6a = QTreeWidgetItem(child6)
        child6a.setText(0, '装药')
        child6b = QTreeWidgetItem(child6)
        child6b.setText(0, '爆破实体')
        child6c = QTreeWidgetItem(child6)
        child6c.setText(0, '底部结构')

        child7 = QTreeWidgetItem(root)
        child7.setText(0, '通用插件')
        child71 = QTreeWidgetItem(child7)
        child71.setText(0, '闭合段生成实体')
        child72 = QTreeWidgetItem(child7)
        child72.setText(0, '外推矿体生成实体')
        child73 = QTreeWidgetItem(child7)
        child73.setText(0, '劈分闭合线圈')
        child74 = QTreeWidgetItem(child7)
        child74.setText(0, '求实体的质心')
        child75 = QTreeWidgetItem(child7)
        child75.setText(0, '内插过渡段')
        child76 = QTreeWidgetItem(child7)
        child76.setText(0, 'XY坐标换')
        child77 = QTreeWidgetItem(child7)
        child77.setText(0, '点击隐藏实体')
        child78 = QTreeWidgetItem(child7)
        child78.setText(0, '将DTM中每个网赋以独有的体号')
        child79 = QTreeWidgetItem(child7)
        child79.setText(0, '将DTM中所有网统一到同一体号')
        child7a = QTreeWidgetItem(child7)
        child7a.setText(0, '强制性连线')
        child7b = QTreeWidgetItem(child7)
        child7b.setText(0, '选择两个点切剖面')
        child7c = QTreeWidgetItem(child7)
        child7c.setText(0, '生成坐标网格体')
        child7d = QTreeWidgetItem(child7)
        child7d.setText(0, '沿线生成打印散点')
        child7e = QTreeWidgetItem(child7)
        child7e.setText(0, '为等高线赋Z值')
        child7f = QTreeWidgetItem(child7)
        child7f.setText(0, '坡顶坡底赋Z值（最短距离）')
        child7g = QTreeWidgetItem(child7)
        child7g.setText(0, '坡顶坡底赋Z值（手工）')
        child7h = QTreeWidgetItem(child7)
        child7h.setText(0, '按长度分割线条')
        child7i = QTreeWidgetItem(child7)
        child7i.setText(0, '生成示坡线')
        child7j = QTreeWidgetItem(child7)
        child7j.setText(0, '生成巷道断面')
        child7k = QTreeWidgetItem(child7)
        child7k.setText(0, '以面积求多边形')
        child7l = QTreeWidgetItem(child7)
        child7l.setText(0, '将2根线在交点断开')

        self.addTopLevelItem(root)

    def addItems(self, parent, itemTexts):
        for itemText in itemTexts:
            _item = QTreeWidgetItem(parent)
            _item.setText(0, itemText)

    def onItemClicked(self, item):
        print(item.text(0))
        print(item.text(1))
        print(item.text(2))
        print(item.text(3))
        if (item.text(1)):
            thread = MyThread(self.port, item)
            thread.start()
            time.sleep(0.5)
            thread.terminate()


class MySurpac:
    # 启动执行文件返回进程pid
    def startProcess(cmd):
        # cmd = "C:/Program Files (x86)/GEOVIA/Surpac/69/nt_i386/bin/surpac2.exe"
        pid = subprocess.Popen(cmd).pid
        return pid

    # 从指定pid获取窗口句柄（通过回调函数）
    def getHwndFromPid(pid):
        def callback(hwnd, hwnds):
            if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
                if found_pid == pid:
                    print(win32gui.GetWindowText(hwnd))
                    hwnds.append(hwnd)
            return True

        hwnds = []
        win32gui.EnumWindows(callback, hwnds)
        return hwnds

    # 从指定名称获取进程的pid数组
    def getPidsFromPName(pname):
        _result = subprocess.Popen("tasklist|findstr " + pname, shell=True, stdout=subprocess.PIPE)
        _lines = _result.stdout.readlines()
        pids = []
        for pid in _lines:
            begin = str(pid).index('surpac2.exe') + 11
            end = str(pid).index('Console')
            pids.append(str(pid)[begin:end].strip())
        return pids

    # 根据pid获取运行端口
    def getPortsFromPid(pid):
        _result = subprocess.Popen("netstat -aon|findstr " + str(pid), shell=True, stdout=subprocess.PIPE)
        _lines = _result.stdout.readlines()
        ports = []
        for port in _lines:
            _port = str(port).replace(' ', '') \
                .replace('.', '') \
                .replace(':', '') \
                .replace("b'", '') \
                .replace("\\r\\n'", '') \
                .replace('TCP0000', '') \
                .replace(str(pid), '') \
                .replace('00000LISTENING', '')
            if len(_port) <= 10:
                ports.append(_port)
        return ports

    # 关闭列出的所有进程id号的进程
    def killProcess(pids):
        for pid in pids:
            _pid = int(pid)
            try:
                os.kill(_pid, signal.SIGTERM)
                print('Process(pid=%s) has be killed' % pid)
            except OSError:
                print('no such process(pid=%s)' % pid)

    # 通过pid获取包含指定窗口特征名的窗口句柄
    def getTheMainWindow(pid, spTitle):
        hwnds = []
        while True:
            hwnds = MySurpac.getHwndFromPid(pid)
            if (len(hwnds) > 0):
                _title = win32gui.GetWindowText((hwnds[0]))
                if (spTitle in _title):
                    break
            time.sleep(1)
        return hwnds[0]

    # 显示窗口
    def showWindow(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_SHOW)

    # 隐含窗口
    def hiddenWindow(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_HIDE)

    # 关闭窗口
    def closeWindow(hwnd):
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE)

    # 设置窗口样式
    def setNoTitleWindow(hwnd):
        ISTYLE = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        win32gui.SetWindowLong(hwnd,
                               ISTYLE &
                               ~win32con.WS_CAPTION &
                               win32con.SWP_NOMOVE &
                               win32con.SWP_NOSIZE)

    # 将一个窗口句柄转化为一个标准Widget
    def convertWndToWidget(hwnd):
        native_wnd = QWindow.fromWinId(hwnd)
        return QWidget.createWindowContainer(native_wnd)


class MyTabWidget(QTabWidget):
    def __init__(self):
        super(MyTabWidget, self).__init__()
        self.setTabsClosable(True)
        self.setDocumentMode(True)
        self.setMovable(True)
        self.tabCloseRequested.connect(self.closeTabItem)

    def createTabItem(self, widget, tabItemTitle):
        tabItem = QWidget()
        layout = QHBoxLayout()
        layout.setSpacing(30)
        layout.setDirection(QHBoxLayout.LeftToRight)
        layout.addWidget(widget)
        tabItem.setLayout(layout)
        self.addTab(tabItem, tabItemTitle)
        self.setCurrentWidget(tabItem)

    def closeTabItem(self, index):
        if self.count() > 1:
            self.removeTab(index)
        else:
            self.close()


class MySocketClient:

    def __init__(self, port: object, encode: object) -> object:
        self.HOST = 'localhost'
        self.BUFSIZ = 1024
        self.PORT = port
        self.ENCODE = encode
        self.ADDR = (self.HOST, port)
        self.tcpCliSock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.tcpCliSock.connect(self.ADDR)

    def sendMsg(self, msg):
        self.tcpCliSock.sendall(msg.encode(self.ENCODE))
        result = self.tcpCliSock.recv(self.BUFSIZ)
        return result

    def closeSocket(self):
        self.tcpCliSock.close()


class MyWebView(QWebEngineView):

    def __init__(self, tabWidget, parent=None):
        super(MyWebView, self).__init__(parent)
        self.tabWidget = tabWidget
        self.settings().setAttribute(QWebEngineSettings.PluginsEnabled, True)  # 支持视频播放
        self.page().windowCloseRequested.connect(self.on_windowCloseRequested)  # 页面关闭请求
        self.page().profile().downloadRequested.connect(self.on_downloadRequested)  # 页面下载请求

    #  支持页面关闭请求
    def on_windowCloseRequested(self):
        the_index = self.mainwindow.tabWidget.currentIndex()
        self.mainwindow.tabWidget.removeTab(the_index)

    #  支持页面下载按钮
    def on_downloadRequested(self, downloadItem):
        if downloadItem.isFinished() == False and downloadItem.state() == 0:
            # 生成文件存储地址
            the_filename = downloadItem.url().fileName()
            if len(the_filename) == 0 or "." not in the_filename:
                cur_time = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
                the_filename = "下载文件" + cur_time + ".xls"
            the_sourceFile = os.path.join(os.getcwd(), the_filename)

            # 下载文件
            # downloadItem.setSavePageFormat(QWebEngineDownloadItem.CompleteHtmlSaveFormat)
            downloadItem.setPath(the_sourceFile)
            downloadItem.accept()
            downloadItem.finished.connect(self.on_downloadfinished)

    #  下载结束触发函数
    def on_downloadfinished(self):
        js_string = '''
        alert("下载成功，请到软件同目录下，查找下载文件！"); 
        '''
        self.page().runJavaScript(js_string)

    # 重载QWebEnginView的createwindow()函数
    def createWindow(self, QWebEnginePage_WebWindowType):
        new_webview = MyWebView(self.tabWidget)
        self.tabWidget.createTabItem(new_webview, '新页面')
        return new_webview
