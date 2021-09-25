#需求：1.设定获取路径、输出路径
#2.修改Pass为Good，并同步板名和日期
#3.可设置抓取频率
#4.设定读取延时

import os
import csv
import sys
import time
from shutil import move
import configparser
#文件对话框
import tkinter.filedialog
#文件监控头文件
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer
from PyQt5.QtWidgets import QApplication,QWidget,QGridLayout,QPushButton,QVBoxLayout,QHBoxLayout
from PyQt5.QtWidgets import QRadioButton,QButtonGroup,QFileDialog,QLineEdit
from PyQt5.QtWidgets import QCheckBox,QLabel,QComboBox,QListWidget,QTextEdit
from PyQt5.QtGui import QFont
from PyQt5.QtCore import QObject,pyqtSlot,pyqtSignal

class FileEventHandler(FileSystemEventHandler,QObject):
    alertsetter = pyqtSignal(str,str)#str1为类型，str2为文件路径
    def on_moved(self, event):
        if event.is_directory:
            print("directory moved from {} to {}".format(event.src_path,event.dest_path))
        else:
            print("file moved from {} to {}".format(event.src_path,event.dest_path))


    def on_created(self, event):
        if event.is_directory:
            print("directory created:{}".format(event.src_path))
        else:
            print("file created:{}".format(event.src_path))
            self.alertsetter.emit("create", event.src_path)


    def on_deleted(self, event):
        if event.is_directory:
            print("directory deleted:{}".format(event.src_path))
        else:
            print("file deleted:{}".format(event.src_path))

    def on_modified(self, event):
        if event.is_directory:
            print("directory modified:{}".format(event.src_path))
        else:
            print("file modified:{}".format(event.src_path))

    def task(self,filename):
        print(filename)
        #具体任务

class Winform(QWidget):
    def __init__(self,parent=None):
        super(Winform,self).__init__(parent)
        self.initUI()

    def initUI(self):
        #初始化计数,记录过板数
        self.bcount=0
        self.jgvar=0
        #初始化变量，初始化路径（置为空）
        self.fromwenjianjia =""
        self.towenjianjia = ""
        #初始化watchdog
        self.observer = Observer()
        self.event_handler = FileEventHandler()
        self.event_handler.alertsetter.connect(self.ObserveHandle)
        #总layout
        vlayout = QVBoxLayout()

        #第一层layout
        hlayout1 = QHBoxLayout()
        label1=QLabel("选择抓取路径：")
        self.lineedit_from = QLineEdit()
        self.pushbutton_changefrom = QPushButton("选取")
        hlayout1.addWidget(label1, 0)
        hlayout1.addWidget(self.lineedit_from, 1)
        hlayout1.addWidget(self.pushbutton_changefrom, 0)
        vlayout.addLayout(hlayout1)
        #第二层layout
        hlayout2 = QHBoxLayout()
        label2=QLabel("选择输出路径：")
        self.lineedit_to = QLineEdit()
        self.pushbutton_changeto = QPushButton("选取")
        hlayout2.addWidget(label2, 0)
        hlayout2.addWidget(self.lineedit_to, 1)
        hlayout2.addWidget(self.pushbutton_changeto, 0)
        vlayout.addLayout(hlayout2)
        self.setLayout(vlayout)
        #状态layout
        hlayoutx = QHBoxLayout()
        self.labelx = QLabel("运行状态：")
        self.labelstatus = QLabel("未运行")
        self.labelx.setFont(QFont("Microsoft YaHei", 12))
        self.labelstatus.setFont(QFont("Microsoft YaHei", 12))
        hlayoutx.addWidget(self.labelx,0)
        hlayoutx.addWidget(self.labelstatus,1)
        vlayout.addLayout(hlayoutx)
        #第三层layout
        self.textedit1=QTextEdit("日志：")
        vlayout.addWidget(self.textedit1)
        #第四层layout
        hlauout4 = QHBoxLayout()
        label4 = QLabel("设定间隔")
        pushbutton_jg=QPushButton("设定")
        self.lineedit_jg=QLineEdit()
        hlauout4.addWidget(label4,0)
        hlauout4.addWidget(self.lineedit_jg, 1)
        hlauout4.addWidget(pushbutton_jg, 0)

        label5 = QLabel("抓取延时：")
        label6 = QLabel("s")
        self.lineedit_jqyy=QLineEdit()
        pushbutton_jqyy = QPushButton("设定")
        hlauout4.addWidget(label5, 0)
        hlauout4.addWidget(self.lineedit_jqyy, 1)
        hlauout4.addWidget(label6, 0)
        hlauout4.addWidget(pushbutton_jqyy, 0)
        vlayout.addLayout(hlauout4)
        #过板
        #第五层layout
        hlayout5=QHBoxLayout()
        self.pushbutton_start=QPushButton("开始")
        self.pushbutton_stop = QPushButton("停止")
        self.pushbutton_stop.setEnabled(0)
        self.pushbutton_start.setMinimumHeight(50)
        self.pushbutton_stop.setMinimumHeight(50)
        self.pushbutton_start.setFont(QFont("Microsoft YaHei", 12))
        self.pushbutton_stop.setFont(QFont("Microsoft YaHei", 12))
        hlayout5.addWidget(self.pushbutton_start,1)
        hlayout5.addWidget(self.pushbutton_stop, 1)
        vlayout.addLayout(hlayout5)
        #信号绑定
        self.pushbutton_changefrom.clicked.connect(self.ChangeFromLocation)
        self.pushbutton_changeto.clicked.connect(self.ChangeToLocation)
        pushbutton_jg.clicked.connect(self.ChangeJg)
        pushbutton_jqyy.clicked.connect(self.ChangeZqYs)
        self.pushbutton_start.clicked.connect(self.StartObserve)
        self.pushbutton_stop.clicked.connect(self.StopObserve)
        self.Initini()
    # 初始化路径信息/读取ini文件
    def Initini(self):
        try:
            self.cfg1 = "config.ini"
            self.conf = configparser.ConfigParser()
            self.conf.read(self.cfg1)
            self.fromwenjianjia = self.conf.get("route", "from")
            self.towenjianjia = self.conf.get("route", "to")
            textjg = self.conf.get("route", "jg")
            self.jg = int(textjg)
            textzqys = self.conf.get("route", "zqys")
            self.zqys = int(textzqys)
        except Exception:
            self.textedit1.append("<font color=\"#FF0000\">路径信息初始化失败，请检查config.ini文件</font>")
        else:
            self.lineedit_from.setText(self.fromwenjianjia)
            self.lineedit_to.setText(self.towenjianjia)
            self.lineedit_jg.setText(textjg)
            self.lineedit_jqyy.setText(textzqys)
    # 修改来源文件夹路径
    def ChangeFromLocation(self):
        fileroute = tkinter.filedialog.askdirectory()
        if (fileroute != ''):
            try:
                wr1=configparser.ConfigParser()
                wr1.add_section("route")
                wr1.set("route", "from", fileroute)
                wr1.set("route", "to", self.towenjianjia)
                wr1.set("route", "jg", str(self.jg))
                wr1.set("route", "zqys",str(self.zqys))
                # write to file
                wr1.write(open(self.cfg1, "w"))
            except Exception:
                self.textedit1.append("<font color=\"#FF0000\">路径更改失败</font>")
            else:
                self.textedit1.append("<font color=\"#0000FF\">抓取文件夹路径已更改为{}</font>".format(fileroute))
                self.lineedit_from.setText(fileroute)
                self.fromwenjianjia=fileroute
        else:
            self.textedit1.append("本次选择未指定路径")
    # 修改输出文件夹路径
    def ChangeToLocation(self):
        fileroute = tkinter.filedialog.askdirectory()
        if (fileroute != ''):
            try:
                wr2 = configparser.ConfigParser()
                wr2.add_section("route")
                wr2.set("route", "from",self.fromwenjianjia)
                wr2.set("route", "to",fileroute)
                wr2.set("route", "jg", str(self.jg))
                wr2.set("route", "zqys",str(self.zqys))
                # write to file
                wr2.write(open(self.cfg1, "w"))
            except Exception:
                self.textedit1.append("<font color=\"#FF0000\">路径更改失败</font>")
            else:
                self.textedit1.append("<font color=\"#0000FF\">输出文件夹路径已更改为{}</font>".format(fileroute))
                self.lineedit_to.setText(fileroute)
                self.towenjianjia = fileroute
        else:
            self.textedit1.append("本次选择未指定路径")
    #修改抓取频率
    def ChangeJg(self):
        text=self.lineedit_jg.text().strip()
        try:
            textnum=int(text)
        except Exception:
            self.textedit1.append("<font color=\"#FF0000\">设置失败，请检查输入是否为规范数字</font>")
        else:
            self.jg = textnum
            #修改ini文件
            try:
                wr1 = configparser.ConfigParser()
                wr1.add_section("route")
                wr1.set("route", "from", self.fromwenjianjia)
                wr1.set("route", "to", self.towenjianjia)
                wr1.set("route", "jg", str(self.jg))
                wr1.set("route", "zqys", str(self.zqys))
                # write to file
                wr1.write(open(self.cfg1, "w"))
            except Exception:
                self.textedit1.append("<font color=\"#FF0000\">写入配置失败，请检查ini文件</font>")
            else:
                self.textedit1.append("<font color=\"#0000FF\">间隔已修改为{}</font>".format(text))

    #修改抓取延时
    def ChangeZqYs(self):
        text = self.lineedit_jqyy.text().strip()
        try:
            textnum = int(text)
        except Exception:
            self.textedit1.append("<font color=\"#FF0000\">设置失败，请检查输入是否为规范数字</font>")
        else:
            self.zqys = textnum
            # 修改ini文件
            try:
                wr1 = configparser.ConfigParser()
                wr1.add_section("route")
                wr1.set("route", "from", self.fromwenjianjia)
                wr1.set("route", "to", self.towenjianjia)
                wr1.set("route", "jg", str(self.jg))
                wr1.set("route", "zqys", str(self.zqys))
                # write to file
                wr1.write(open(self.cfg1, "w"))
            except Exception:
                self.textedit1.append("<font color=\"#FF0000\">写入配置失败，请检查ini文件</font>")
            else:
                self.textedit1.append("<font color=\"#0000FF\">抓取延时已修改为{}s</font>".format(text))
    #修改表格文件
    def ChangeCsv(self,filepath):
        #延迟抓取
        if(self.zqys!=0):
            time.sleep(self.zqys)
        try:
            with open(filepath, 'r') as f:
                reader = csv.reader(f)
                writetext = list(reader)
                for i in range(9, 13):  # 修改头部Pass均改为Good
                    writetext[i][1] = "Good"
                writetext[14][1] = "Good"#Test status改为Good
                # 修改其他部分，同步日期和板SN
                for j in range(16, len(writetext)):
                    writetext[j][13] = "Good"
                # 拆分路径
                (Fpath, tempfilename) = os.path.split(filepath)  # 拆分路径和文件名
                (Fname, extension) = os.path.splitext(tempfilename)  # 拆分真正文件名和后缀
                #Fname = PHNJ13650NSC1285A0_20210910000645_Pass
                Fname = Fname[0:-4]+"Good"
                # 保存
                with open(self.towenjianjia + "\\" + Fname + ".csv", 'w', encoding='utf-8', newline='') as f2:
                    writer = csv.writer(f2)
                    for item in writetext:
                        writer.writerow(item)
                    print("处理完毕")
        except Exception:
            self.textedit1.append("<font color=\"#FF0000\">发现新增{}，出现异常，处理失败</font>".format(tempfilename))
            return -1
        else:
            self.textedit1.append("<font color=\"#0000FF\">发现新增{}并完成处理</font>".format(tempfilename))
            return 0

    #开始监测文件夹
    def StartObserve(self):
        self.pushbutton_start.setEnabled(0)
        self.pushbutton_stop.setEnabled(1)
        self.pushbutton_changefrom.setEnabled(0)
        self.pushbutton_changeto.setEnabled(0)
        if(self.observer.isAlive()):
            self.observer.stop()
            self.textedit1.append("<font color=\"#0000FF\">监测到预先有进程正在运行，已自动关闭</font>")
        self.observer = Observer()
        self.observer.schedule(self.event_handler, self.fromwenjianjia, True)#监视fromwenjianjia路径
        self.observer.start()
        #self.textedit1.append("<font color=\"#0000FF\">监视执行中</font>")
        self.labelstatus.setText("<font color=\"#22DD22\">监视执行中</font>")
    def StopObserve(self):
        self.pushbutton_start.setEnabled(1)
        self.pushbutton_stop.setEnabled(0)
        self.pushbutton_changefrom.setEnabled(1)
        self.pushbutton_changeto.setEnabled(1)
        if (self.observer.isAlive()):
            self.observer.stop()
            self.observer.unschedule_all()
            self.observer.join()
            #self.textedit1.append("<font color=\"#FF0000\">监视已停止</font>")
            self.labelstatus.setText("<font color=\"#FF0000\">监视已停止</font>")
        else:
            self.textedit1.append("<font color=\"#0000FF\">线程未开启，无操作</font>")
    #文件监测发出信号后，接收并处理
    #type为类型，path为路径
    def ObserveHandle(self,type,path):
        if(type=="create"):
            # 拆分路径
            (Fpath, tempfilename) = os.path.split(path)  # 拆分路径和文件名
            (Fname, extension) = os.path.splitext(tempfilename)  # 拆分真正文件名和后缀
            if(Fname[-4:]=="Pass"):#如果是pass则处理，其他则直接转移
                #统计板数+1
                self.bcount=self.bcount+1
                #为0则默认执行，不为0则验证间隔
                if(self.jg==0):
                    self.ChangeCsv(path)
                #判断间隔
                elif(self.bccount%self.jg!=0):
                    self.ChangeCsv(path)
                else:
                    time.sleep(self.zqys)
                    try:
                        move(path, self.towenjianjia)#剪切走
                    except Exception:
                        self.textedit1.append("<font color=\"#FF0000\">尝试剪切失败</font>")
                    else:
                        self.textedit1.append("<font color=\"#0000FF\">文件{}移动成功</font>".format(tempfilename))
            else:
                # 拆分路径
                (Fpath, tempfilename) = os.path.split(path)  # 拆分路径和文件名
                (Fname, extension) = os.path.splitext(tempfilename)  # 拆分真正文件名和后缀
                time.sleep(self.zqys)
                try:
                    move(path, self.towenjianjia)#剪切走
                except Exception:
                    self.textedit1.append("<font color=\"#FF0000\">尝试剪切失败</font>")
                else:
                    self.textedit1.append("<font color=\"#0000FF\">文件{}移动成功</font>".format(tempfilename))
        else:
            pass

    
if __name__ == "__main__":
    app = QApplication(sys.argv)
    form = Winform()
    form.setWindowTitle("PASS监控软件")
    form.setMinimumWidth(600)
    form.setFont(QFont("Microsoft YaHei",11))
    form.show()
    sys.exit(app.exec_())