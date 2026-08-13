import ctypes
from tkinter import *
from tkinter.ttk import *
from PIL import Image,ImageTk

class SelectWindow(Toplevel):

    def __init__(self):
        super().__init__()
        #告诉操作系统使用程序自身的dpi适配
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        #获取屏幕的缩放因子
        ScaleFactor=ctypes.windll.shcore.GetScaleFactorForDevice(0)
        #设置程序缩放
        self.tk.call('tk', 'scaling', ScaleFactor/60)
        self.geometry("550x700+500+300")
        self.title("样式选择")
        self.resizable(0,0)
        self.iconbitmap("resources\\icon\\icon.ico")
        self.selVar = StringVar()
        self.selVar.set("1")
        self.setupUI()
        self.protocol("WM_DELETE_WINDOW",self.CancelDo)

    def setupUI(self):
        style = Style()
        style.configure("title.TLabel",font=("微软雅黑",10))
        style.configure("stitle.TLabel",font=("微软雅黑",7))
        
        self.titlelabel = Label(self,text="请选择要启用的搜索按钮样式",style="title.TLabel")
        self.titlelabel.pack(anchor="n",pady=10)

        self.lbTitle1 = Label(self,text="[ 选项 ]",style="stitle.TLabel")
        self.lbTitle1.place(x=28,y=60)
        self.lbTitle2 = Label(self,text="[ 样式预览图（点击前 | 点击后） ]",style="stitle.TLabel")
        self.lbTitle2.place(x=110,y=60)

        self.i1 = Image.open(r"resources\sbimg\1.png")
        self.img1 = ImageTk.PhotoImage(self.i1)
        self.i2 = Image.open(r"resources\sbimg\2.png")
        self.img2 = ImageTk.PhotoImage(self.i2)
        self.i3 = Image.open(r"resources\sbimg\3.png")
        self.img3 = ImageTk.PhotoImage(self.i3)
        self.i4 = Image.open(r"resources\sbimg\4.png")
        self.img4 = ImageTk.PhotoImage(self.i4)
        self.i5 = Image.open(r"resources\sbimg\5.png")
        self.img5 = ImageTk.PhotoImage(self.i5)

        self.rb1 = Radiobutton(self,variable=self.selVar,value="1",image=self.img1)
        self.rb1.place(x=50,y=100)
        self.rb2 = Radiobutton(self,variable=self.selVar,value="2",image=self.img2)
        self.rb2.place(x=50,y=200)
        self.rb3 = Radiobutton(self,variable=self.selVar,value="3",image=self.img3)
        self.rb3.place(x=50,y=300)
        self.rb4 = Radiobutton(self,variable=self.selVar,value="4",image=self.img4)
        self.rb4.place(x=50,y=400)
        self.rb5 = Radiobutton(self,variable=self.selVar,value="5",image=self.img5)
        self.rb5.place(x=50,y=500)

        self.okBtn = Button(self,text="确定",command=self.OkDo)
        self.okBtn.place(x=113,y=610)
        self.cancelBtn = Button(self,text="取消",command=self.CancelDo)
        self.cancelBtn.place(x=290,y=610)

        
    def OkDo(self):
        self.destroy()

    def CancelDo(self):
        self.selVar.set("0")
        self.destroy()
        


if __name__ == "__main__":
    sw = SelectWindow()
