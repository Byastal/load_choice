import sys
import os
import time
from PyQt5 import QtWidgets
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from readExcel import load_choice
global time_consum,count,url,filename,flag
import subprocess
filename =""
flag =True
# url ="http://127.0.0.1:8000/api/financial/real_time_data"
url ="http://124.71.113.79:80/api/financial/real_time_data"
class WorkThread(QThread):
    # 初始化线程
    def __int__(self):
        super(WorkThread, self).__init__()
        
    #线程运行函数
    def run(self):
        global url,filename,flag,count,time_consum
        flag,count,time_consum = True,0,0
        while flag:
            pre = time.time()
            try:
                response = load_choice(filename,url)
                time_consum+=time.time()-pre
                if response//100!=2:
                    flag = False
                    break
                print(response)
                count+=1
                if flag==True:time.sleep(1)
            except BaseException as e:
                print(e)
                flag=False
        

class FirstMainWindow(QMainWindow):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        #目录路径
        self.firmware_dir = []
        self.setWindowTitle('金融数据抓取工具')
        ###### 创建界面 ######
        self.centralwidget = QWidget()
        self.setCentralWidget(self.centralwidget)
        self.Layout = QVBoxLayout(self.centralwidget)

        # 设置顶部三个按钮
        self.topwidget = QWidget()
        self.Layout.addWidget(self.topwidget)
        self.buttonLayout = QHBoxLayout(self.topwidget)

        self.pushButton1 = QPushButton()
        self.pushButton1.setText("打开文件")
        self.buttonLayout.addWidget(self.pushButton1)

        self.pushButton2 = QPushButton()
        self.pushButton2.setText("更新数据")
        self.buttonLayout.addWidget(self.pushButton2)

        self.pushButton3 = QPushButton()
        self.pushButton3.setText("打开提示框")
        self.buttonLayout.addWidget(self.pushButton3)

        ###### 三个按钮事件 ######
        self.pushButton1.clicked.connect(self.on_pushButton1_clicked)
        self.pushButton2.clicked.connect(self.on_pushButton2_clicked)
        self.pushButton3.clicked.connect(self.on_pushButton3_clicked)

    # 打开子界面
    windowList = []
    def open_update_win(self):
        
        new = plotwindows()
        self.windowList.append(new)   ##注：没有这句，是不打开另一个主界面的！
        new.show()
        self.close()

    # 打开文件
    def on_pushButton1_clicked(self):
        global filename
        filename,fileType = QFileDialog.getOpenFileName(self, "选取文件", "./", 
        "Text Files(*.xlsm)")
        print(filename)
        print(fileType)
        if fileType=="Text Files(*.xlsm)":
            try:
                ret = subprocess.call('start wps office '+filename, shell=True)
                print(ret)
            # command = ["start",filename]
            # ret = subprocess.run(command)
            # os.system("start "+filename)
                self.firmware_dir.append(filename)
            except BaseException as e:
                QMessageBox.information(self, "提示", "找不到excel软件打开文件！")

    # 按钮二：打开对话框
    def on_pushButton2_clicked(self):
        try:
            global filename
            # filename =self.firmware_dir[-1]
            # file_split = self.firmware_dir[-1].split("/")
            # print(filename)
            # if file_split[-1] != "跨期监控表格.xlsm":
            if filename.split("/")[-1] != "跨期监控表格.xlsm":
                QMessageBox.information(self, "提示", "请选择正确的xlsm模板文件！")
            self.open_update_win()
            
        except BaseException as e:
            QMessageBox.information(self, "提示", "更新数据失败，请查看是否有打开xlsm模板！")
            print(e)

    # 按钮三：打开提示框
    def on_pushButton3_clicked(self):
        QMessageBox.information(self, "提示", "请先打开文件，再执行更新数据！")


class plotwindows(QtWidgets.QWidget):
    def __init__(self):
        super(plotwindows,self).__init__()
        self.setWindowTitle('更新情况')
        # 线程
        self.workThread = WorkThread()
        self.workThread.start()
        # 页面布局
        layout = QFormLayout()
        self.edita3 = QLineEdit()
        self.edita4 = QLineEdit()
        self.edita5 = QLineEdit()
        layout.addRow("更新数量", self.edita3)
        layout.addRow("总计耗时", self.edita4)
        self.setLayout(layout)
        # 定时器
        # 这里不要写成一个时间类，不然在子函数不好做回调
        # 因为之前封成一个类，在子方法update中循环调用提示框，导致系统卡死
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update)
        self.timer.start(100)
        
    
    def update(self):
        global count,time_consum,flag
        # print(flag)
        if flag==True:
            self.edita3.setText(str(count))
            self.edita4.setText(str(time_consum)[:5]+" s")
        else:
            self.timer.stop()
            QMessageBox.information(self,"提示","服务器更新数据失败")
            self.close()
            
    windowList = []
    def closeEvent(self, event):
        global flag
        flag = False
        the_window = FirstMainWindow()
        self.windowList.append(the_window)  ##注：没有这句，是不打开另一个主界面的！
        the_window.show()
        if not self.workThread.isRunning():
            return
        self.workThread.quit()      # 退出
        self.workThread.wait()      # 回收资源
        event.accept()
        

if __name__ == "__main__":
    app = QApplication(sys.argv)
    the_mainwindow = FirstMainWindow()
    the_mainwindow.show()
    sys.exit(app.exec_())

