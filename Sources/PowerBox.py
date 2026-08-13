# coding:utf-8
# PowerBox Full (GUI)
# Ver. 6.8.0 (build.714-2026)
# Copyright © 2026 luluandpipi. All right reserves.

from __future__ import print_function
from tkinter import *
from tkinter.ttk import *
import ctypes,sys,win32api,win32con,os
import subprocess as sp
import platform
from time import sleep
from SearchButtonSelect import SelectWindow
from win32com.client import Dispatch

winver_full = platform.version()
winver = winver_full.split(".")

choice1 = [
        "请选择",
        "请选择",
        "文件资源管理器-回退到 Win10 样式",
        "文件资源管理器-回退到 Win7 样式",
        "文件资源管理器-回到 Win11 样式",
        "右键菜单-回退到 Win10 样式",
        "右键菜单-回到 Win11 样式",
        "任务栏-回退到 Win10 样式",
        "任务栏-回到 Win11 样式",
        "开始菜单-回退到 Win10 样式",
        "开始菜单-回到 Win11 样式",
        "设置-回退到 Win10 样式",
        "设置-回到 Win11 样式"
        ]

choice2 = ["请选择",
           "请选择",
           "启用小组件的新入口（适用于 22518 及以上版本，22533 及以上版本已默认启用）",
           "关闭小组件的新入口（适用于 22518 及以上版本）",
           "启用任务管理器新样式（适用于 22538 及以上版本，22557 及以上版本已默认启用）",
           "关闭任务管理器新样式（适用于 22538 及以上版本）",
           "启用桌面贴纸（适用于 22563 及以上版本）",
           "关闭桌面贴纸（适用于 22563 及以上版本）",
           "启用文件资源管理器标签页（适用于 22572、22579.100、22581.100 及以上版本，25136 及以上版本已默认启用）",
           "关闭文件资源管理器标签页（适用于 22572、22579.100、22581.100 及以上版本）",
           "启用新版桌面搜索框（适用于 25120 及以上版本）",
           "关闭新版桌面搜索框（适用于 25120 及以上版本）",
           "启用任务栏的新搜索按钮（适用于 25136 及以上版本）",
           "关闭任务栏的新搜索按钮（适用于 25136 及以上版本）",
           "启用任务栏多样搜索按钮（适用于 25158 及以上版本）",
           "关闭任务栏多样搜索按钮（适用于 25158 及以上版本）",
           "启用新版任务栏启动动画（适用于 25179 版本）",
           "关闭新版任务栏启动动画（适用于 25179 版本）",
           "启用无启动动画的任务栏（适用于 25182 及以上版本）",
           "关闭无启动动画的任务栏（适用于 25182 及以上版本）",
           "启用新版设置菜单栏点击动画（适用于 25188 及以上版本）",
           "关闭新版设置菜单栏点击动画（适用于 25188 及以上版本）"
           ]

choice3 = ["请选择",
           "请选择",
           "打开 IE 浏览器",
           "打开/关闭桌面搜索框"
           ]
 
class Application:
    def __init__(self):
        self.root = Tk()
        #告诉操作系统使用程序自身的dpi适配
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        #获取屏幕的缩放因子
        ScaleFactor=ctypes.windll.shcore.GetScaleFactorForDevice(0)
        #设置程序缩放
        self.root.tk.call('tk', 'scaling', ScaleFactor/60)
        self.root.geometry("800x550+500+300")
        self.root.title("PowerBox 全能版")
        self.root.iconbitmap("resources\\icon\\icon.ico")
        self.setupUI()
        self.maximize_notice()

        self.flag_win7 = 0
        
    def setupUI(self):
        style = Style()
        style.configure("a.TLabel",font=("微软雅黑",14))
        style.configure("b.TLabel",font=("微软雅黑",10))
        style.configure("c.TLabel",font=("微软雅黑",7))
        style.configure("a.TButton",font=("微软雅黑",11),width=13)
        
        self.titlelabel = Label(self.root,text="PowerBox",style="a.TLabel")
        self.titlelabel.pack(anchor="n",pady=10)

        self.verlabel = Label(self.root,text="全能版 6.8.0",style="c.TLabel")
        self.verlabel.place(x=490,y=28)

        self.authorlabel = Label(self.root,text="作者：lulu&pipi",style="c.TLabel")
        self.authorlabel.place(x=330,y=60)

        self.boxlabel_a = Label(self.root,text="Win10 样式回退",style="b.TLabel")
        self.boxlabel_a.place(x=35,y=130)

        self.boxlabel_b = Label(self.root,text="Win11 Dev 功能尝鲜",style="b.TLabel")
        self.boxlabel_b.place(x=295,y=130)

        self.boxlabel_c = Label(self.root,text="小工具",style="b.TLabel")
        self.boxlabel_c.place(x=615,y=130)

        self.var1 = StringVar()
        self.var1.set(choice1[0])

        self.var2 = StringVar()
        self.var2.set(choice2[0])

        self.var3 = StringVar()
        self.var3.set(choice3[0])
        
        self.menu1 = OptionMenu(self.root,self.var1,*choice1) #*号有一个解包的功能，如果没有 * 号，就把整个列表当做一个选项了
        self.menu1.place(x=75,y=170)

        self.menu2 = OptionMenu(self.root,self.var2,*choice2) #*号有一个解包的功能，如果没有 * 号，就把整个列表当做一个选项了
        self.menu2.place(x=345,y=170)

        self.menu3 = OptionMenu(self.root,self.var3,*choice3) #*号有一个解包的功能，如果没有 * 号，就把整个列表当做一个选项了
        self.menu3.place(x=605,y=170)

        self.startbutton = Button(self.root,text="开始",style="a.TButton",command=self.detect)
        self.startbutton.place(x=291,y=300)
 
    def detect(self):
        if self.var1.get() != choice1[0] and self.var2.get() != choice2[0] or self.var1.get() != choice1[0] and self.var3.get() != choice3[0] or self.var2.get() != choice2[0]\
           and self.var3.get() != choice3[0]:
            win32api.MessageBox(0,"每次只能选择一个功能哦","提示",win32con.MB_ICONINFORMATION)
        elif self.var1.get() == choice1[0]:
            self.detect2()
        else:
            if self.var1.get() == choice1[2]:
                if winver[2] >= "22621":
                    if win32api.MessageBox(0,"请注意，在 Win11 22621 及以上版本中进行此操作将使右键菜单一并回退到 Win10 样式。确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                        self.win10styleexplorer()
                else:
                    if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                        self.win10styleexplorer()
            elif self.var1.get() == choice1[3]:
                if win32api.MessageBox(0,"确定要开始执行吗？进行此操作将同时使右键菜单回退到 Win10 样式。","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.flag_win7 = True
                    self.win7styleexplorer()
            elif self.var1.get() == choice1[4]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.win11styleexplorer()
            elif self.var1.get() == choice1[5]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.win10stylecontext()
            elif self.var1.get() == choice1[6]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.win11stylecontext()
            elif self.var1.get() == choice1[7]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.win10styletaskbar()
            elif self.var1.get() == choice1[8]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.win11styletaskbar()
            elif self.var1.get() == choice1[9]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.win10stylestart()
            elif self.var1.get() == choice1[10]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.win11stylestart()
            elif self.var1.get() == choice1[11]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.win10stylesettings()
            elif self.var1.get() == choice1[12]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.win11stylesettings()


    def detect2(self):
        if self.var2.get() == choice2[0]:
            self.detect3()
        else:
            if self.var2.get() == choice2[2]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.enableAugmentedEntryPoint()
            elif self.var2.get() == choice2[3]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.disableAugmentedEntryPoint()
            elif self.var2.get() == choice2[4]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.enableNewTaskMgr()
            elif self.var2.get() == choice2[5]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.disableNewTaskMgr()
            elif self.var2.get() == choice2[6]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.enableStickers()
            elif self.var2.get() == choice2[7]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.disableStickers()
            elif self.var2.get() == choice2[8]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.enableTabExplorer()
            elif self.var2.get() == choice2[9]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.disableTabExplorer()
            elif self.var2.get() == choice2[10]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.enableDesktopSearch()
            elif self.var2.get() == choice2[11]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.disableDesktopSearch()
            elif self.var2.get() == choice2[12]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.enableNewSearchButton()
            elif self.var2.get() == choice2[13]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.disableNewSearchButton()
            elif self.var2.get() == choice2[14]:
                self.enableColorfulSearchButton()
            elif self.var2.get() == choice2[15]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.disableColorfulSearchButton()
            elif self.var2.get() == choice2[16]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.enableNewTbAnimation()
            elif self.var2.get() == choice2[17]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.disableNewTbAnimation()
            elif self.var2.get() == choice2[18]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.enableNoAnimTb()
            elif self.var2.get() == choice2[19]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.disableNoAnimTb()
            elif self.var2.get() == choice2[20]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.enableNewSetSbAnim()
            elif self.var2.get() == choice2[21]:
                if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                    self.disableNewSetSbAnim()

    def detect3(self):
        if self.var3.get() == choice3[0]:
            win32api.MessageBox(0,"请先选择一个功能","提示",win32con.MB_ICONINFORMATION)
        else:
            if self.var3.get() == choice3[2]:
                self.launchIE()
            elif self.var3.get() == choice3[3]:
                self.launchKillDSB()

    def finishAction(self,actStyle):
        if actStyle == 1:
            sp.run('taskkill /f /im explorer.exe',shell=True)
            sp.run('start explorer.exe',shell=True)
            win32api.MessageBox(0,"操作已成功完成","提示",win32con.MB_ICONINFORMATION)

        elif actStyle == 2:
            sp.run('taskkill /f /im explorer.exe',shell=True)
            sp.run('start explorer.exe',shell=True)

        elif actStyle == 3:
            if win32api.MessageBox(0,"操作完成。需要重启以启用更改，是否立即重启？","提示",win32con.MB_YESNO) == win32con.IDYES:
                sp.run(r'@shutdown /r /t 0',shell=True)
            
    def win10styleexplorer(self):
        if winver[2] >= "22621":
            sp.run(r"takeown /f C:\Windows\System32\Windows.UI.FileExplorer.dll >nul 2>&1 && icacls C:\Windows\System32\Windows.UI.FileExplorer.dll /grant administrators:F /c /q >nul 2>&1",shell=True)
            sp.run(r"takeown /f C:\Windows\System32\UIRibbon.dll >nul 2>&1 && icacls C:\Windows\System32\UIRibbon.dll /grant administrators:F /c /q >nul 2>&1",shell=True)
            sp.run(r"ren C:\Windows\System32\Windows.UI.FileExplorer.dll Windows.UI.FileExplorer.dll.backup",shell=True)
            if not self.flag_win7:
                self.finishAction(1)
        else:
            sp.run(r'reg add "HKLM\Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked" /t REG_SZ /v {e2bf9676-5f8f-435c-97eb-11607a5bedf7} /f',shell=True)
            if not self.flag_win7:
                self.finishAction(1)

    def win7styleexplorer(self):
        self.win10styleexplorer()
        sp.run(r"ren C:\Windows\System32\UIRibbon.dll UIRibbon.dll.backup",shell=True)
        self.finishAction(1)

    def win11styleexplorer(self):
        if winver[2] >= "22621":
            sp.run(r"ren C:\Windows\System32\Windows.UI.FileExplorer.dll.backup Windows.UI.FileExplorer.dll",shell=True)
            if not self.flag_win7:
                self.finishAction(1)
        else:
            sp.run(r'reg delete "HKLM\Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked" /f 2>nul',shell=True)
            if not self.flag_win7:
                self.finishAction(1)
        if self.flag_win7:
            sp.run(r"ren C:\Windows\System32\UIRibbon.dll.backup UIRibbon.dll",shell=True)
            self.flag_win7 = False
            self.finishAction(1)


    def win10stylecontext(self):
        sp.run(r'reg add "HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32" /f',shell=True)
        self.finishAction(1)

    def win11stylecontext(self):
        sp.run(r'reg delete "HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}" /f 2>nul',shell=True)
        self.finishAction(1)

    def win10styletaskbar(self):
        sp.run(r'reg add "HKLM\Software\Microsoft\Windows\CurrentVersion\Shell\Update\Packages" /t REG_DWORD /v UndockingDisabled /d 1 /f',shell=True)
        self.finishAction(2)
        win32api.MessageBox(0,"操作完成，请在打开的窗口中启用想要在任务栏中显示的图标","提示",win32con.MB_ICONINFORMATION)
        sleep(1)
        sp.run('start shell:::{05d7b0f4-2121-4eff-bf6b-ed3f69b894d9}\SystemIcons',shell=True)

    def win11styletaskbar(self):
        sp.run(r'reg delete "HKLM\Software\Microsoft\Windows\CurrentVersion\Shell\Update\Packages" /v UndockingDisabled /f 2>nul',shell=True)
        self.finishAction(1)

    def win10stylestart(self):
        sp.run(r"taskkill /f /im startmenuexperiencehost.exe",shell=True)
        sp.run(r"takeown /f C:\Windows\systemapps\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy /r /d y >nul 2>&1 && icacls C:\windows\systemapps\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\ /grant administrators:F /t /c /q >nul 2>&1",shell=True)
        sp.run(r"takeown /f C:\Windows\SystemResources\Windows.UI.ShellCommon\ /r /d y >nul 2>&1 && icacls C:\Windows\SystemResources\Windows.UI.ShellCommon\ /grant administrators:F /t /c /q >nul 2>&1",shell=True)
        sp.run(r"takeown /f C:\Windows\ShellExperiences /r /d y >nul 2>&1 && icacls C:\Windows\ShellExperiences\ /grant administrators:F /t /c /q >nul 2>&1",shell=True)

        sp.run(r"taskkill /f /im startmenuexperiencehost.exe",shell=True)
        sp.run(r"ren C:\Windows\SystemApps\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\StartMenuExperienceHost.exe StartMenuExperienceHost.exe.backup",shell=True)
        sp.run(r"ren C:\Windows\SystemApps\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\StartDocked.dll StartDocked.dll.backup",shell=True)
        sp.run(r"ren C:\Windows\SystemApps\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\pris\resources.zh-CN.pri resources.zh-CN.pri.backup",shell=True)
        sp.run(r"ren C:\Windows\SystemResources\Windows.UI.ShellCommon\Windows.UI.ShellCommon.pri Windows.UI.ShellCommon.pri.backup",shell=True)
        sp.run(r"ren C:\Windows\SystemResources\Windows.UI.ShellCommon\pris\Windows.UI.ShellCommon.zh-CN.pri Windows.UI.ShellCommon.zh-CN.pri.backup",shell=True)

        sp.run(r"taskkill /f /im startmenuexperiencehost.exe",shell=True)
        sp.run(r'xcopy "resources\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\StartMenuExperienceHost.exe" "C:\windows\systemapps\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy"',shell=True)
        sp.run(r'xcopy "resources\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\StartDocked.dll" "C:\windows\systemapps\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy"',shell=True)
        sp.run(r'xcopy "resources\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\pris\resources.zh-CN.pri" "C:\windows\systemapps\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\pris"',shell=True)
        sp.run(r'xcopy "resources\ShellCommon\Windows.UI.ShellCommon.pri" "C:\Windows\SystemResources\Windows.UI.ShellCommon"',shell=True)
        sp.run(r'xcopy "resources\ShellCommon\Windows.UI.ShellCommon.zh-CN.pri" "C:\Windows\SystemResources\Windows.UI.ShellCommon\pris"',shell=True)
        
        sp.run(r'reg add "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /t REG_DWORD /v Start_ShowClassicMode /d 1 /f',shell=True)
        sp.run(r"taskkill /f /im startmenuexperiencehost.exe",shell=True)

        if win32api.MessageBox(0,"操作完成。是否将任务栏居左对齐？","提示",win32con.MB_YESNO) == win32con.IDYES:
            sp.run(r'reg add "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /t REG_DWORD /v TaskbarAl /d 0 /f',shell=True)

    def win11stylestart(self):
        if not os.path.exists(r"C:\Windows\SystemApps\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\StartDocked.dll.backup"):
            win32api.MessageBox(0,"你的开始菜单已经是 Win11 样式了哦！","提示",win32con.MB_ICONINFORMATION)
        else:
            sp.run(r"taskkill /f /im startmenuexperiencehost.exe",shell=True)
            sp.run(r"del /f /q C:\Windows\SystemApps\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\StartMenuExperienceHost.exe",shell=True)
            sp.run(r"del /f /q C:\Windows\SystemApps\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\StartDocked.dll",shell=True)
            sp.run(r"del /f /q C:\Windows\SystemApps\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\pris\resources.zh-CN.pri",shell=True)
            sp.run(r"del /f /q C:\Windows\SystemResources\Windows.UI.ShellCommon\Windows.UI.ShellCommon.pri",shell=True)
            sp.run(r"del /f /q C:\Windows\SystemResources\Windows.UI.ShellCommon\pris\Windows.UI.ShellCommon.zh-CN.pri",shell=True)

            sp.run(r"taskkill /f /im startmenuexperiencehost.exe",shell=True)
            sp.run(r'ren C:\Windows\SystemApps\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\StartMenuExperienceHost.exe.backup StartMenuExperienceHost.exe',shell=True)
            sp.run(r'ren C:\Windows\SystemApps\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\StartDocked.dll.backup StartDocked.dll',shell=True)
            sp.run(r'ren C:\Windows\SystemApps\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\pris\resources.zh-CN.pri.backup resources.zh-CN.pri',shell=True)
            sp.run(r'ren C:\Windows\SystemResources\Windows.UI.ShellCommon\Windows.UI.ShellCommon.pri.backup Windows.UI.ShellCommon.pri',shell=True)
            sp.run(r'ren C:\Windows\SystemResources\Windows.UI.ShellCommon\pris\Windows.UI.ShellCommon.zh-CN.pri.backup Windows.UI.ShellCommon.zh-CN.pri',shell=True)

            sp.run(r'reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v Start_ShowClassicMode /f',shell=True)
            sp.run(r"taskkill /f /im startmenuexperiencehost.exe",shell=True)

            if win32api.MessageBox(0,"操作完成。是否将任务栏居中对齐？","提示",win32con.MB_YESNO) == win32con.IDYES:
                sp.run(r'reg add "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /t REG_DWORD /v TaskbarAl /d 1 /f',shell=True)

    def win10stylesettings(self):
        sp.run(r"taskkill /f /im SystemSettings.exe",shell=True)
        sp.run(r"takeown /f C:\Windows\ImmersiveControlPanel\ /r /d y >nul 2>&1 && icacls C:\Windows\ImmersiveControlPanel\ /grant administrators:F /t /c /q >nul 2>&1",shell=True)
        sp.run(r"takeown /f C:\Windows\SystemResources\Windows.UI.SettingsAppThreshold\ /r /d y >nul 2>&1 && icacls C:\Windows\SystemResources\Windows.UI.SettingsAppThreshold\ /grant administrators:F /t /c /q >nul 2>&1",shell=True)

        sp.run(r"ren C:\Windows\ImmersiveControlPanel ImmersiveControlPanel_backup",shell=True)
        sp.run(r"ren C:\Windows\SystemResources\Windows.UI.SettingsAppThreshold\Windows.UI.SettingsAppThreshold.pri Windows.UI.SettingsAppThreshold.pri.backup",shell=True)
        sp.run(r"ren C:\Windows\SystemResources\Windows.UI.SettingsAppThreshold\SystemSettings SystemSettings_backup",shell=True)
        sp.run(r"ren C:\Windows\SystemResources\Windows.UI.SettingsAppThreshold\pris pris_backup",shell=True)
        
        os.mkdir("C:\Windows\ImmersiveControlPanel")
        sp.run(r'xcopy /S "resources\ImmersiveControlPanel" "C:\Windows\ImmersiveControlPanel"',shell=True)
        sp.run(r'xcopy /S "resources\Windows.UI.SettingsAppThreshold" "C:\Windows\SystemResources\Windows.UI.SettingsAppThreshold"',shell=True)

        if win32api.MessageBox(0,"操作完成。是否立即打开设置？","提示",win32con.MB_YESNO) == win32con.IDYES:
            sp.run("start ms-settings:",shell=True)

    def win11stylesettings(self):
        if not os.path.exists(r"C:\Windows\ImmersiveControlPanel_backup"):
            win32api.MessageBox(0,"你的设置已经是 Win11 样式了哦！","提示",win32con.MB_ICONINFORMATION)
        else:
            sp.run(r"taskkill /f /im SystemSettings.exe",shell=True)

            sp.run(r"rmdir /S /Q C:\Windows\ImmersiveControlPanel",shell=True)
            sp.run(r"del /f /q C:\Windows\SystemResources\Windows.UI.SettingsAppThreshold\Windows.UI.SettingsAppThreshold.pri",shell=True)
            sp.run(r"rmdir /S /Q C:\Windows\SystemResources\Windows.UI.SettingsAppThreshold\SystemSettings",shell=True)
            sp.run(r"rmdir /S /Q C:\Windows\SystemResources\Windows.UI.SettingsAppThreshold\pris",shell=True)
            
            sp.run(r"ren C:\Windows\ImmersiveControlPanel_backup ImmersiveControlPanel",shell=True)
            sp.run(r"ren C:\Windows\SystemResources\Windows.UI.SettingsAppThreshold\Windows.UI.SettingsAppThreshold.pri.backup Windows.UI.SettingsAppThreshold.pri",shell=True)
            sp.run(r"ren C:\Windows\SystemResources\Windows.UI.SettingsAppThreshold\SystemSettings_backup SystemSettings",shell=True)
            sp.run(r"ren C:\Windows\SystemResources\Windows.UI.SettingsAppThreshold\pris_backup pris",shell=True)

            if win32api.MessageBox(0,"操作完成。是否立即打开设置？","提示",win32con.MB_YESNO) == win32con.IDYES:
                sp.run("start ms-settings:",shell=True)

    def FeatureMgmt(self,FID,inst):
        if inst:
            cString = 'reg add \"HKLM\\SYSTEM\\CurrentControlSet\\Control\\FeatureManagement\\Overrides\\4\\'+FID+'\" '
            sp.run(cString + '/t REG_DWORD /v EnabledState /d 2 /f',shell=True)
            sp.run(cString + '/t REG_DWORD /v EnabledStateOptions /d 1 /f',shell=True)
            sp.run(cString + '/t REG_DWORD /v Variant /d 0 /f',shell=True)
            sp.run(cString + '/t REG_DWORD /v VariantPayload /d 0 /f',shell=True)
            sp.run(cString + '/t REG_DWORD /v VariantPayloadKind /d 0 /f',shell=True)
        else:
            cString = 'reg delete \"HKLM\\SYSTEM\\CurrentControlSet\\Control\\FeatureManagement\\Overrides\\4\\'+FID+'\" '
            sp.run(cString + '/f 2>nul',shell=True)

    def sbtnSelect(self):
        self.sw = SelectWindow()
        self.root.wait_window(self.sw)

    def enableAugmentedEntryPoint(self):
        self.FeatureMgmt("2212930700",1)
        self.FeatureMgmt("233221772",1)
        self.FeatureMgmt("3466913933",1)
        self.FeatureMgmt("837070477",1)

        self.finishAction(3)

    def disableAugmentedEntryPoint(self):
        self.FeatureMgmt("2212930700",0)
        self.FeatureMgmt("233221772",0)
        self.FeatureMgmt("3466913933",0)
        self.FeatureMgmt("837070477",0)

        self.finishAction(3)

    def enableNewTaskMgr(self):
        self.FeatureMgmt("2534314125",1)
        self.FeatureMgmt("644126861",1)
        self.FeatureMgmt("96123020",1)

        self.finishAction(3)

    def disableNewTaskMgr(self):
        self.FeatureMgmt("2534314125",0)
        self.FeatureMgmt("644126861",0)
        self.FeatureMgmt("96123020",0)

        self.finishAction(3)

    def enableStickers(self):
        sp.run(r'reg add "HKLM\SOFTWARE\Microsoft\PolicyManager\current\device\Stickers" /t REG_DWORD /v EnableStickers /d 1 /f',shell=True)

        self.finishAction(2)
        win32api.MessageBox(0,"操作完成，在桌面上右键单击，选择“添加或编辑贴纸”即可","提示",win32con.MB_ICONINFORMATION)

    def disableStickers(self):
        sp.run(r'reg delete "HKLM\SOFTWARE\Microsoft\PolicyManager\current\device\Stickers" /f',shell=True)

        self.finishAction(1)

    def enableTabExplorer(self):
        self.FeatureMgmt("1351149197",1)

        self.finishAction(3)

    def disableTabExplorer(self):
        self.FeatureMgmt("1351149197",0)

        self.finishAction(3)

    def enableDesktopSearch(self):
        self.FeatureMgmt("3067508877",1)

        self.finishAction(3)

    def disableDesktopSearch(self):
        self.FeatureMgmt("3067508877",0)

        self.finishAction(3)

    def enableNewSearchButton(self):
        self.FeatureMgmt("3272906381",1)

        self.finishAction(3)

    def disableNewSearchButton(self):
        self.FeatureMgmt("3272906381",0)

        self.finishAction(3)

    def enableColorfulSearchButton(self):
        self.sbtnSelect()
        if self.sw.selVar.get() == "0":
            pass
        else:
            if win32api.MessageBox(0,"确定要开始执行吗？","提示",win32con.MB_OKCANCEL) == win32con.IDOK:
                self.FeatureMgmt("3255588492",1)
                cString = r'reg add "HKLM\SYSTEM\CurrentControlSet\Control\FeatureManagement\Overrides\4\3255588492" /t REG_DWORD /v Variant /d ' + self.sw.selVar.get() + ' /f'
                sp.run(cString,shell=True)
                self.finishAction(3)

    def disableColorfulSearchButton(self):
        self.FeatureMgmt("3255588492",0)

        self.finishAction(3)

    def enableNewTbAnimation(self):
        self.FeatureMgmt("3255588492",1)

        self.finishAction(3)

    def disableNewTbAnimation(self):
        self.FeatureMgmt("3255588492",0)

        self.finishAction(3)

    def enableNoAnimTb(self):
        self.FeatureMgmt("628245132",1)

        self.finishAction(3)

    def disableNoAnimTb(self):
        self.FeatureMgmt("628245132",0)

        self.finishAction(3)

    def enableNewSetSbAnim(self):
        self.FeatureMgmt("2416629389",1)

        sleep(0.2)
        if win32api.MessageBox(0,"操作完成。需要重启以启用更改，是否立即重启？\n重启后打开设置，点击左侧菜单栏中的任意一项即可看到新的图标动画","提示",\
                               win32con.MB_YESNO) == win32con.IDYES:
            sp.run(r'@shutdown /r /t 0',shell=True)

    def disableNewSetSbAnim(self):
        self.FeatureMgmt("2416629389",0)

        self.finishAction(3)

    def launchIE(self):
        sp.run("taskkill /f /im iexplore.exe",shell=True)
        IE = Dispatch("InternetExplorer.Application")
        IE.Navigate("https://cn.bing.com")
        IE.Toolbar = 1
        IE.visible = True

    def launchKillDSB(self):
        winver_full = platform.version()
        winver = winver_full.split(".")
        if winver[2] < "25120":
            win32api.MessageBox(0,"此功能仅支持 25120 及以上版本","提示",win32con.MB_ICONINFORMATION)
        else:
            if not os.path.exists(r'.\resources\config'):
                os.mkdir(r'.\resources\config')
            if not os.path.exists(r'.\resources\config\flags'):
                os.mkdir(r'.\resources\config\flags')
            if not os.path.exists(r'.\resources\config\flags\DSBLaunched.flag'):
                open(r'.\resources\config\flags\DSBLaunched.flag',"w")
                sp.run(r"start shell:AppsFolder\MicrosoftWindows.Client.CBS_cw5n1h2txyewy!DesktopSearchBoxCentennial",shell=True)
            else:
                sp.run("taskkill /f /im DesktopSearchBoxWin32Exe.exe",shell=True)
                os.remove(r'.\resources\config\flags\DSBLaunched.flag')

    def maximize_notice(self):
        if self.root.state() == "zoomed":
            self.root.state("normal")
            win32api.MessageBox(0,"不要把我弄得那么胖，会爆掉的！","警告",win32con.MB_ICONWARNING)
        if self.root.winfo_width() != 800 or self.root.winfo_height() != 550:
            self.root.geometry("800x550")
            #win32api.MessageBox(0,"不要扯来扯去，很难受的！","警告",win32con.MB_ICONWARNING)
        self.root.after(800,self.maximize_notice)
                
    def run(self):
        self.root.mainloop()

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False
if is_admin():
    app = Application()
    app.run()
else:
    if sys.version_info[0] == 3:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    else:#in python2.x
        ctypes.windll.shell32.ShellExecuteW(None, u"runas", unicode(sys.executable), unicode(__file__), None, 1)
