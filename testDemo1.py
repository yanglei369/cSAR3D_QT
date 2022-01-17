#-*- coding:utf-8 -*-
import win32gui
import win32api
import win32con
import tkinter as tk
from PIL import Image,ImageTk
from tkinter import PhotoImage
# from tkinter import *
from tkinter import messagebox,scrolledtext,Button,Label,SUNKEN,END
from win32gui import PyMakeBuffer, SendMessage,PostMessage, PyGetBufferAddressAndLen, PyGetString, PySetMemory, PyGetMemory,GetParent
import threading
import pyautogui
from time import sleep
import pywin32_testutil
import array
import time
import os
import shutil





class App(object):
    """app初始化"""
    def __init__(self,root):
        self.root = root
        self.consoleText = ""
        self.cameraFlag = None
        self.childHwndList = None
        self.top_hwnd_list = []
        self.main_window()

    def GetHwndListByTitle(self,allTitle):
        try:
            # 每次只要解析标题，都要将cSAR3D置为最前面
            top_hwnd_list = []
            # self.validParentHwnd = None
            # childHwndList = []
            for h, t in self.hwnd_titles.items():
                if allTitle in t or allTitle == t:
                    top_hwnd_list.append(h)
            return top_hwnd_list
        except Exception as e:
            self.print_log(e)
            return None

    def GetChildHwndListByFatherHwnd(self,fatherHwnd):
            childList = []
            try:
                win32gui.EnumChildWindows(fatherHwnd, lambda hwnd, param: param.append(hwnd), childList)
                return childList
            except Exception as e:
                self.print_log(e)
                return None

    def GetValidHwnd(self,HwndList):
        for singleHwnd in HwndList:
            childList = []
            try:
                win32gui.EnumChildWindows(singleHwnd, lambda hwnd, param: param.append(hwnd), childList)
                if len(childList) == 0:
                    self.writeLogToLogFile("\n{}为无效句柄".format(singleHwnd))
                else:
                    return singleHwnd
            except Exception as e:
                self.print_log(e)
                return None

    def setFGWindowByHwnd(self,validHwnd):
        win32gui.SetForegroundWindow(validHwnd)
        win32gui.ShowWindow(validHwnd, win32con.SW_NORMAL)

    def SkipCamera(self):
        sleep(1)
        # if not self.refreshWindowsAndParseTopHwnd("cSAR3DApplication"):
        #     return
        camera = 0
        while True:
            camera += 1
            if camera <= 4:
                if not self.refreshWindowsAndParseTopHwnd("Camera"):
                    break
                if (self.cameraFlag):
                    cameraHwnd = self.getHwndListByTitle("Skip Calibration")
                    # self.print_log("cameraHwnd is {}".format(cameraHwnd))
                    if len(cameraHwnd) == 1:
                        cameraPos = self.getPosByHwnd(cameraHwnd[0])
                    elif len(cameraHwnd) == 0:
                        break
                    else:
                        # self.print_log("获取相机标签数量异常!")
                        self.writeLogToLogFile("\n获取相机标签数量异常!")
                        self.finishFlag = True
                        break
                    # 模拟点击camera
                    # sleep(1)
                    # self.print_log(cameraPos)
                    self.clickXyCenter(cameraPos)
                    sleep(2)
            else:
                break

    def print_log(self,mytext):
        #print(mytext)
        #print("{}\n".format(mytext))
        #print(type(mytext))
        self.text.insert(END, "{}\n".format(mytext))
        #保持焦点在行尾
        self.text.see(END)
        self.text.update()

    #获取所有句柄和标题组成的item
    def get_all_hwnd_dir(self, hwnd, mouse):
        self.hwnd_titles = {}
        if (win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd)):
            self.hwnd_titles.update({hwnd: win32gui.GetWindowText(hwnd)})

    #获取所有窗口句柄标题
    def scanWindowGetAllTitleName(self):
        try:
            #EnumWindows并不枚举子窗口, EnumWindows对屏幕上所有的顶层窗口依次进行枚举，将每个顶层窗口的窗口句柄依次传递给回调函数。
            win32gui.EnumWindows(self.get_all_hwnd_dir, 0)
            return self.hwnd_titles
        except Exception as e:
            return None


    def enumChildHwndListByTopHwnd(self,top_hwnd_list):
        #self.print_log("进入1这里....")
        for singleParentHwnd in top_hwnd_list:
            childList = []
            try:
                win32gui.EnumChildWindows(singleParentHwnd, lambda hwnd, param: param.append(hwnd), childList)
                if len(childList) == 0:
                    #self.top_hwnd_list.remove(singleParentHwnd)
                    #self.print_log("{}无效".format(singleParentHwnd))
                    self.writeLogToLogFile("\n{}为无效句柄".format(singleParentHwnd))
                else:
                    #self.print_log(childHwndList)
                    #self.print_log("进入2这里")
                    self.childHwndList = childList
                    self.validParentHwnd = singleParentHwnd
                    break
            except Exception as e:
                self.print_log(e)

    def parseAndPrepositionTopHwnd(self,topTitle):
        try:
            #每次只要解析标题，都要将cSAR3D置为最前面
            top_hwnd_list = []
            self.validParentHwnd = None
            childHwndList = []
            for h, t in self.hwnd_titles.items():
                if topTitle in t or topTitle == t:
                    top_hwnd_list.append(h)

            if topTitle == "Camera":
                if len(top_hwnd_list) == 0:
                    self.print_log("此次未获取camera.")
                    self.writeLogToLogFile("\n此次未获取camera.")
                    self.cameraFlag = False
                    return True
                if len(top_hwnd_list) == 1:
                    self.enumChildHwndListByTopHwnd(top_hwnd_list)
                    if len(childHwndList) == 0:
                        # self.print_log("此次未获取camera..")
                        self.writeLogToLogFile("\n此次未获取camera..")
                        self.cameraFlag = False
                    return True

            if topTitle == "cSAR3DApplication":
                if len(top_hwnd_list) == 0:
                    self.print_log("请确认是否已经打开cSAR3D软件!")
                    self.writeLogToLogFile("\n请确认是否已经打开cSAR3D软件!")
                    self.finishFlag = True
                    return False
                if len(top_hwnd_list) == 1:
                    self.enumChildHwndListByTopHwnd(top_hwnd_list)
                    win32gui.SetForegroundWindow(top_hwnd_list[0])
                    win32gui.ShowWindow(top_hwnd_list[0], win32con.SW_NORMAL)
                    return True

            if len(top_hwnd_list) > 1:
                self.enumChildHwndListByTopHwnd(top_hwnd_list)
                if len(childHwndList) == 0 and topTitle == "Camera":
                    #self.print_log("此次未获取camera...")
                    self.writeLogToLogFile("\n此次未获取camera...")
                    #self.writeLogToLogFile("\n请确认是否已选择测试工程或窗口未在屏幕内 ?")
                    #a = tk.messagebox.askyesno('提示', '请确认是否已选择测试工程或窗口未在屏幕内 ?')
                    #if not a:
                        #return False
                    self.cameraFlag = False
                    return True

                elif len(childHwndList) == 0 and topTitle == "cSAR3DApplication":
                    #self.print_log("cSAR3D枚举异常!")
                    self.writeLogToLogFile("\ncSAR3D枚举异常!")
                    self.finishFlag = True
                    return False


                try:
                    if topTitle == "cSAR3DApplication" and self.validParentHwnd != None:
                            win32gui.SetForegroundWindow(self.validParentHwnd)
                            win32gui.ShowWindow(self.validParentHwnd, win32con.SW_NORMAL)
                except:
                    #self.print_log("未获取到有效窗口, 设置窗口前置失败!")
                    self.writeLogToLogFile("\n未获取到有效窗口, 设置窗口前置失败!")
                    self.finishFlag = True
                    return False
                return True

        except Exception as e:
            self.finishFlag = True
            self.print_log(e)
            return False

    def getTitleNameByHwnd(self,jbid):
        # 获取标题
        title = win32gui.GetWindowText(jbid)
        # print(title)
        return title

    def getHwndListByTitle(self,titleName):
        HWND = []
        try:
            if self.childHwndList:
                #self.print_log(self.childHwndList)
                #self.print_log("现在获取的标签是{}".format(titleName))
                for curChildHwnd in self.childHwndList:
                    t = self.getTitleNameByHwnd(curChildHwnd)
                    if titleName in t or titleName == t:
                        HWND.append(curChildHwnd)
            if not HWND:
                self.writeLogToLogFile("\n标题{}解析失败!".format(titleName))
                #self.print_log("标题{}解析失败!".format(titleName))
                return None
            return HWND
        except Exception as e:
            self.print_log(e)
            self.finishFlag = True
            return

    def getPosByHwnd(self, hwnd):
        try:
            l, t, r, b = win32gui.GetWindowRect(hwnd)
            # help(pyautogui.doubleClick)
            x = l + (r - l) / 2
            y = t + (b - t) / 2
            return (x,y)
            # self.print_log({"x": x, "y": y})
            # self.print_log("rect: " + str((l, t, r, b)))
        except Exception as e:
            self.print_log(e)
            return

    def getOffsetPosByHwnd(self, hwnd, titie):
        try:
            if titie == "Measurement":
                l, t, r, b = win32gui.GetWindowRect(hwnd)
                #self.print_log("l:{},t:{},r:{},b:{}".format(l,t,r,b))
                #self.writeLogToLogFile("\nl:{},t:{},r:{},b:{}".format(l,t,r,b))
                # help(pyautogui.doubleClick)
                x = l + 126
                y = t - 23 + 11
                return (x,y)
                # self.print_log({"x": x, "y": y})
                # self.print_log("rect: " + str((l, t, r, b)))
            if titie == "Step Project":
                l, t, r, b = win32gui.GetWindowRect(hwnd)
                # help(pyautogui.doubleClick)
                x = l + 40
                y = t - 23 + 11
                return (x, y)
            if titie == "Data Analysis":
                l, t, r, b = win32gui.GetWindowRect(hwnd)
                # help(pyautogui.doubleClick)
                x = l + 210
                y = t - 23 + 11
                return (x, y)
        except Exception as e:
            self.print_log(e)
            return

    def clickXyCenter(self,itemXY):
        try:
            # itemXY = self.address_data_7.get()
            #self.print_log("现在点击的坐标是{}".format(itemXY))
            x = int(itemXY[0])
            y = int(itemXY[1])
            # self.print_log(x)
            # self.print_log(y)
            # try:
            #     for parentHwnd in self.cur_SendMessage_HWND_List:
            #         # print(type(parentHwnd))
            #         # if str.isdigit(parentHwnd):
            #         str = ""
            #         if parentHwnd != str and parentHwnd.isdigit():
            #             win32gui.SetForegroundWindow(parentHwnd)
            # except Exception as e:
            #     self.print_log(e)
            #     self.print_log("请先点击解析父句柄!")
            #     return

            pyautogui.moveTo(x, y,duration=0.1)
            pyautogui.click(x, y, button="left")
            #pyautogui.hotkey("ctrl", "a")
            #pyautogui.hotkey("ctrl", "c")

            #data = self.root.clipboard_get()
            #print(data)
            # data = pyperclip.paste()
            #self.print_log("获取到的内容为{}.".format(data))
        except Exception as e:
            self.print_log(e)
            #self.print_log("请检查输入格式: x,y格式必须以x,y形式输入,逗号也不能使用中文！")
            return

    def parseStartHwnd(self):
        try:
            startHwndList = None
            self.leftStartPos = None
            self.rightStartPos = None

            startHwndList = self.getHwndListByTitle("Start")
            #self.print_log("Start标签的句柄有{}".format(startHwndList))
            if len(startHwndList) == 1:
                self.leftStartPos = self.getPosByHwnd(startHwndList[0])
                return True
            elif len(startHwndList) == 2:
                # 获取坐标
                startList = []
                for everyStartHwnd in startHwndList:
                    startList.append(self.getPosByHwnd(everyStartHwnd))
                a = startList[0][0]
                #self.print_log("a是{}".format(a))
                b = startList[1][0]
                #self.print_log("b是{}".format(b))
                if a > b:
                    self.leftStartPos = startList[1]
                    self.rightStartPos = startList[0]
                if a < b:
                    self.leftStartPos = startList[0]
                    self.rightStartPos = startList[1]
                return True
            else:
                #self.print_log("获取start标签数量异常,请确认是否已切换到Start窗口!")
                self.writeLogToLogFile("\n获取start标签数量异常,请确认是否已切换到Start窗口!")
                return False
        except Exception as e:
            self.print_log(e)
            return False

    def refreshWindowsAndParseTopHwnd(self,topTitle):
        try:
            self.getAllTitleName()
            if (self.parseAndPrepositionTopHwnd(topTitle)):
                return True
            else:
                return False
        except Exception as e:
            self.print_log(e)
            self.print_log("222")
            return False

    def sendContent(self,consoleHwnd):
        # hwnd = int(self.address_data_9.get())
        # try:
        #     win32gui.SetForegroundWindow(hwnd)
        # except Exception as e:
        #     self.print_log(e)
        #     self.print_log("句柄{}无法设置为前置显示！".format(hwnd))
        content = "clear...\n"
        # 发生消息
        # 如果向uart assistant发送消息，则不能进行encode编码， 直接进行发送即可， 向notepad++则需要进行utf-8编码
        win32api.SendMessage(consoleHwnd, win32con.WM_SETTEXT, None, content.encode('utf-8'))
        #win32api.SendMessage(hwnd, win32con.WM_SETTEXT, None, content)

    def getTextBoxContentByHwnd(self,consoleHwnd):
        try:
            #self.clearText()
            #self.text.update()
            text = ""
            # 缓冲区只保留待查内容的长度，打印多了会打印出不必要的乱码
            length = SendMessage(consoleHwnd, win32con.WM_GETTEXTLENGTH) * 2
            if length >= 200000:
                #self.print_log("控制台内容操过100M!")
                self.writeLogToLogFile("\n控制台内容操过100M!")
                return
            if length == 0:
                #self.print_log("控制台无内容!")
                self.writeLogToLogFile("\n控制台无内容!")
                return
            structStr = ""
            # 从指定列表随机生成多个数， 我这里是随机生成一个数，因为生成的是列表，要取出这个数，需要取索引[0]
            # list = [0,1,2,3,4,5,6]
            # k = (random.sample(list, 1))[0]
            for i in range(length):
                structStr += ("\{}".format(0))

            # ASCII也可以 字符串编码为字节
            z = str(eval(repr(structStr).replace('\\\\', '\\'))).encode('gb2312')

            # 手动构造内存空间   字符串编码为字节
            test_data = pywin32_testutil.str2bytes(str(z,encoding = "utf-8"))   #utf-8和ASCII也得到一样的结果
            c = array.array("b",z)
            addr, buflen = c.buffer_info()
            buf = win32gui.PyGetMemory(addr,buflen)
            #buf = PyMakeBuffer(length)
            SendMessage(consoleHwnd, win32con.WM_GETTEXT, length, buf)

            try:
                address, length1 = PyGetBufferAddressAndLen(buf)
            except ValueError:
                #self.print_log("获取文本框地址出错！")
                self.writeLogToLogFile("\n获取文本框地址出错!")
                return
            text = PyGetString(address, length1)

            realLength = int(length / 2) + 1
            text = text[0:realLength]
            #self.print_log("\n获取到的文本框内容为: \n{}".format(text))
            #self.writeLogToLogFile("\n获取到的文本框内容为: \n{}".format(text))
            buf.release()
            del buf
            del c
            return text
            # python显式的垃圾回收器
            # gc.collect()

        except Exception as e:
            self.print_log(e)
            return

    # def getTextBoxContentByHwnd_thread(self,consoleHwnd):
    #     try:
    #         #time.sleep(0.1)
    #         t1 = threading.Thread(target=self.getTextBoxContentByHwnd,args=(consoleHwnd,))
    #         # 守护线程
    #         t1.setDaemon(True)
    #         # 启动
    #         t1.start()
    #         # self.print_log("start thread successful!")
    #     except Exception as e:
    #         t1.join()

    def collectConsoleLog(self,consoleHwnd):
        text = self.getTextBoxContentByHwnd(consoleHwnd)
        return text

    def runRobotArm(self):
        num = 0
        while(num < 5):
            self.print_log("机械臂工作中...")
            sleep(2)
            num += 1
        self.print_log("机械臂工作完成!")

    def collectLogAndAnalyse(self,consoleHwnd):
        num = 0
        while True:
            num += 1
            #self.print_log("\n\n测试轮数:{}轮".format(num))
            #self.writeLogToLogFile("\n\n测试轮数:{}轮".format(num))
            sleep(8)
            if self.pauseFlag:
                break
            self.consoleText = ""
            self.consoleText = self.collectConsoleLog(consoleHwnd[0])
            #self.print_log("\n文本长度为{}".format(len(self.consoleText)))
            # 分析控制台文本
            if "Finalizing" in self.consoleText:
                if self.pauseFlag:
                    break
                self.runRobotArm()
                self.print_log("test success!")
                self.writeLogToLogFile("\n\n测试轮数:{}轮".format(num))
                sleep(2)
                self.writeLogToLogFile("\n\ntest success: \n{}\n\n*************end*************".format(self.consoleText))
                self.root.state('normal')
                self.finishFlag = True
                tk.messagebox.showinfo("提示", "   当前用例测试完成 !  ")
                break
            elif len(self.consoleText) > 100 and "Finalizing" not in self.consoleText:
                sleep(1)
                self.runRobotArm()
                # 模拟点击右侧start按钮
                #self.print_log("第{}次重复点击start".format(num))
                if num <= 15:
                    self.clickXyCenter(self.rightStartPos)
                else:
                    self.print_log("测试流程非标准流程!")
                    break
                #点击后，需要等待一下测试，再进行日志收集
                #sleep(1)
            else:
                if num < 10:
                    continue
                if consoleHwnd[0] and num > 10:
                    self.print_log("获取测试结果异常!")
                    break

    def writeLogToLogFile(self,log):
        with open(r"C:\result\mylog.log", 'ab') as fp:
            fp.write(log.encode('utf-8'))
            #sleep(1)
            fp.close()

    def initLogFile(self):
        #开启工具界面时间打印
        self.printTime()
        #本地日志文件初始化
        if (os.path.exists(r"C:\result")):
            shutil.rmtree(r"C:\result")
        #本地创建目录用于保存结果文件
        os.makedirs(r"C:\result")
        if (os.path.exists(r"C:\result")):
            pass
        else:
            self.writeLogToLogFile("\n创建C盘临时目录失败!")
            self.finishFlag = True
            return
        #创建本地日志文件
        now = time.localtime()
        nowt = time.strftime("%Y-%m-%d  %H_%M_%S", now)  # 这一步就是对时间进行格式化
        if (os.path.exists(r"C:\result\mylog.log")):
            os.remove(r"C:\result\mylog.log")
            self.writeLogToLogFile("*************start*************\n{}".format(nowt))
        else:
            self.writeLogToLogFile("*************start*************\n{}".format(nowt))

        # 状态初始化： 软件最小化，避免影响cSAR句柄抓取
        self.root.state('icon')         #窗口有3中状态，iconic：最小化；normal：正常显示；zoomed：最大化。
        self.cameraFlag = True
        self.pauseFlag = False

    def printTimeThread(self):
        while True:
            if not self.finishFlag:
                now = time.localtime()
                nowt = time.strftime("%Y-%m-%d  %H_%M_%S", now)
                self.print_log(nowt)
                sleep(6)
            else:
                break

    def printTime(self):
        # 时间线程
        self.timeThread = threading.Thread(target=self.printTimeThread)
        # 守护线程
        self.timeThread.setDaemon(False)
        # 启动
        self.timeThread.start()

    def startTest(self):
        # 打印开始测试信息
        self.print_log("\nstarting... \nplease wait...\n")
        # 扫描窗口获取所有窗口标题
        allTitleName = self.scanWindowGetAllTitleName()
        if allTitleName:
            # 根据软件标题解析出对应句柄列表
            fatherHwndList = self.GetHwndListByTitle("cSAR3DApplication", allTitleName)
            if fatherHwndList is None:
                self.print_log("请确认是否已经打开cSAR3D软件!")
                self.writeLogToLogFile("\n请确认是否已经打开cSAR3D软件!")
                # self.finishFlag = True
                # return False
                return
            else:
                # 解析句柄列表获取有效句柄
                validHwnd = self.GetValidHwnd(fatherHwndList)  # 肯定只有1个是有效的。

                # 根据句柄，前置窗口
                self.setFGWindowByHwnd(validHwnd)

                # 根据父句柄获取子句柄列表
                childHwndList = self.GetChildHwndListByFatherHwnd(validHwnd)
                if childHwndList is None:
                    self.print_log("获取cSAR3D软件子项失败!")
                    self.writeLogToLogFile("\n获取cSAR3D软件子项失败!")
                    return
                else:
                    # 解析句柄列表获取有效句柄
                    self.GetValidHwnd(childHwndList)

                    # 刷新所有窗口，并解析出父句柄
                    if not self.refreshWindowsAndParseTopHwnd("cSAR3DApplication"):
                        # self.print_log("请确认是否打开cSAR3D软件！")
                        return
                    # 先切换到Measurement选项卡
                    measurementHwnd = self.getHwndListByTitle("Measurement")
                    # self.print_log("measurement句柄为{}".format(measurementHwnd))
                    if len(measurementHwnd) == 1:
                        measurementPos = self.getOffsetPosByHwnd(measurementHwnd[0], "Measurement")
                        # self.print_log("measurementPos坐标为{}".format(measurementPos))
                    else:
                        self.writeLogToLogFile("\n获取measurement标签数量异常!")
                        self.finishFlag = True
                        # self.print_log("获取measurement标签数量异常!")
                        return

                    sleep(1)
                    # 模拟点击Measurement
                    self.clickXyCenter(measurementPos)

                    # 刷新所有窗口，并解析出父句柄
                    if not self.refreshWindowsAndParseTopHwnd("cSAR3DApplication"):
                        return

                    # 解析出左右start句柄
                    if not self.parseStartHwnd():
                        return

                    # self.print_log(self.leftStartPos)
                    # self.print_log(self.rightStartPos)

                    # 获取console句柄
                    consoleHwnd = self.getHwndListByTitle("ScintillaConsole")
                    if len(consoleHwnd) != 1:
                        # self.print_log("请打开控制台窗口后，重新测试!")
                        self.writeLogToLogFile("\n请打开控制台窗口后，重新测试!")
                        self.finishFlag = True
                        return
                    if len(self.consoleText) > 100 and "Finalizing" not in self.consoleText \
                            and "Stopping measurements" not in self.consoleText:
                        # self.print_log("进入....{}\n".format(self.consoleText))
                        # 模拟点击start
                        sleep(1)
                        # self.print_log("进入...")
                        self.clickXyCenter(self.leftStartPos)
                        sleep(1)
                        if not self.refreshWindowsAndParseTopHwnd("cSAR3DApplication"):
                            return
                        if not self.parseStartHwnd():
                            return
                        self.collectLogAndAnalyse(consoleHwnd)
                        self.sendContent(consoleHwnd[0])
                        return

                        # 清理console窗口
                        self.sendContent(consoleHwnd[0])

                        # 模拟点击start
                        sleep(1)
                        self.clickXyCenter(self.leftStartPos)

                        # 模拟跳过相机弹窗
                        self.SkipCamera()

                        # 刷新，获取最新的start坐标
                        if not self.refreshWindowsAndParseTopHwnd("cSAR3DApplication"):
                            return

                        # 解析出左右start句柄
                        if not self.parseStartHwnd():
                            return
                        # self.print_log("第二次左{}".format(self.leftStartPos))
                        # self.print_log("第二次右{}".format(self.rightStartPos))

                        # 模拟点击右侧start按钮
                        self.clickXyCenter(self.rightStartPos)

                        # 开启循环监测并收集日志
                        self.collectLogAndAnalyse(consoleHwnd)
                    else:
                        self.print_log("\n获取标题名失败!")
                        self.writeLogToLogFile("\n获取标题名失败!")
                        return

    def run_rpa(self):
        try:
            #程序初始化
            self.initLogFile()
            #开始测试
            self.startTest()
        except Exception as e:
            self.finishFlag = True
            self.print_log(e)



    def app_start(self):
        try:
            self.finishFlag = False
            # 主线程开始
            self.mainThread = threading.Thread(target=self.run_rpa)
            # 守护线程
            self.mainThread.setDaemon(False)
            # 启动
            self.mainThread.start()
        except Exception as e:
            self.print_log(e)
            self.finishFlag = True
            self.mainThread.join()
            self.timeThread.join()

    def stop_rpa(self):
        # 最小化
        #self.root.state('icon')
        self.pauseFlag = True
        self.finishFlag = True
        # 刷新，获取最新的start坐标
        if not self.refreshWindowsAndParseTopHwnd("cSAR3DApplication"):
            return
        try:
            self.consoleText = ""
        except Exception as e:
            self.print_log(e)

        stopHwndList = self.getHwndListByTitle("Stop")
        if len(stopHwndList) == 1:
            stopPos = self.getPosByHwnd(stopHwndList[0])
        else:
            #self.print_log("获取stop标签数量异常!")
            self.writeLogToLogFile("\n获取stop标签数量异常!")
            self.finishFlag = True
            return

        # 模拟点击界面stop
        self.clickXyCenter(stopPos)
        self.print_log("stop success!")
        tk.messagebox.showinfo("提示", "   测试已暂停 !  ")

    def clearText(self):
        self.text.delete('1.0','end')

    def app_stop(self):
        try:
            # 主线程开始
            stopThread = threading.Thread(target=self.stop_rpa)
            # 守护线程
            stopThread.setDaemon(False)
            # 启动
            stopThread.start()
            # 打印开始测试信息
            self.print_log("\nstoping... \nplease wait...\n")

        except Exception as e:
            self.print_log(e)
            stopThread.join()

    # def get_image(self,filename):
    #     im = Image.open(filename)
    #     return ImageTk.PhotoImage(im)

    def setIcon(self):
        from ek import img
        import base64
        tmp = open("tmp2.ico", "wb+")
        tmp.write(base64.b64decode(img))  # 写入到临时文件中
        tmp.close()
        self.root.iconbitmap("tmp2.ico")  # 设置图标
        os.remove("tmp2.ico")

    def get_image(self):
        from png import img
        import base64
        tmp = open("tmp1.ico", "wb+")
        tmp.write(base64.b64decode(img))  # 写入到临时文件中
        tmp.close()
        #self.root.iconbitmap("tmp.ico")  # 设置图标
        im = Image.open("tmp1.ico")
        image = ImageTk.PhotoImage(im)
        os.remove("tmp1.ico")
        return image


    def main_window(self):
        self.root.geometry("650x700")
        self.root.resizable(0, 0)
        #self.root.resizable(False, False)  # 固定窗体
        # self.root.maxsize(600,500)
        # self.root.minsize(600,500)
        self.root.title("essun-grabtool")
        canvas_root = tk.Canvas(self.root,width=650,height=700)
        bg_photo = self.get_image()
        # Label(self.root, image = bg_photo).pack()
        self.setIcon()
        # image = PhotoImage(file="E.png")
        # self.root.iconphoto(False, image)
        #self.root.iconbitmap(r"D:\work\program\language\python\code\RPA\app2.ico")
        self.root.config(bg="blue")
        canvas_root.create_image(325,55,image=bg_photo)    #325 350是中心锚点位置
        canvas_root.pack(anchor='nw')
        startButton = Button(self.root, text="start", font=('新宋体', 15),width=20, bg="LightBlue4", command=self.app_start)
        startButton.place(anchor='sw', relx=0.02, rely=0.08)

        stopButton = Button(self.root, text="stop", font=('新宋体', 15),width=20, bg="LightBlue4", command=self.app_stop)
        stopButton.place(anchor='se', relx=0.98, rely=0.08)

        # 清空日志
        Button(self.root, text="日志 :",bg="LightBlue4").place(anchor='sw', relx=0.02, rely=0.15, height=25)
        delButton = Button(self.root, text="清空日志", bg="LightBlue4", anchor='w', command=self.clearText)
        delButton.place(anchor='se', relx=0.98, rely=0.15, height=25)

        # 日志
        self.text = scrolledtext.ScrolledText(self.root, cursor='arrow', bg='LightCyan4', \
                                              bd=2, fg='blue', relief=SUNKEN)
        self.text.place(anchor='nw', relx=0, rely=0.16, relwidth=0.987, relheight=0.83)
        self.root.mainloop()

if __name__ == '__main__':
    win = tk.Tk()
    App_win = App(win)