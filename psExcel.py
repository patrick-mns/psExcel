from tkinter import *
import tkinter as tk
from tkinter import filedialog as fd
import time
import os 
from psExcelCore import psExcelCore
import webbrowser
import uuid
from datetime import datetime
import sys 

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)
    
def updateDepositLabel(label, txt):
    label.set(txt)

def select_file(typeFile, message, label):
    filetypes = (
        (f'{typeFile}', f'{typeFile}'),
    )
    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='C:\\'+  os.path.abspath(__file__),
        filetypes=filetypes)

    if filename:
        updateDepositLabel(label, filename)
    else:
        None

def select_dir(message, label):
    filename = fd.askdirectory(
        title='Select a directory',
        initialdir='C:\\'+  os.path.abspath(__file__))
    if filename:
        updateDepositLabel(label, filename)
    else:
        None

def openDir(dir):
    path = dir
    path = os.path.realpath(path)
    os.startfile(path)

def errorMessage(window, message):
    window.title("Error")
    Label(window, text=message, background='#2e2e2e', fg='white', font= ("Verdana", "10"), pady=(50)).pack()

def createWindow():
    newWindow = Toplevel(root)
    newWindow.grab_set() 
    newWindow.title("...")
    wx = root.winfo_x() + 100
    wy = root.winfo_y() + 400 - 300
    newWindow.geometry(f"300x100+{wx}+{wy}")
    newWindow.configure(background='#2e2e2e')
    newWindow.iconbitmap(iconfile)
    newWindow.resizable(False, False)
    return(newWindow)

def data():
    def on_closing(pdir):
        try:
            openDir(pdir)
            root.quit
        except:
            openDir(labelTextDir.get())
            root.quit

    newWindow = createWindow()

    if labelTextJson.get() == '' or labelTextPSD.get() == '' or labelTextDir.get() == '' :
        errorMessage(newWindow, 'Invalid import or dir')
        return
    
    pdir = f"{labelTextDir.get()}\\{datetime.today().strftime('%Y-%m-%d') + ' ' + str(uuid.uuid4())}"
    def sendData():
        newWindow.protocol("WM_DELETE_WINDOW", lambda:[newWindow.destroy, on_closing(pdir)])
        newSend = psExcelCore(labelTextJson.get(), labelTextPSD.get(), pdir)
        n = newSend.start()
        if n == True:
            time.sleep(5)
            openDir(pdir)
            newWindow.destroy()
        else:
            newWindow.title("Error")
            newWindow.protocol("WM_DELETE_WINDOW", newWindow.destroy)
            updateDepositLabel(startnofic, n)
            newWindow.after(2000, newWindow.destroy)
            return
            
    startnofic = StringVar()
    startnofic.set('processing')
    Label(newWindow, textvariable=startnofic, background='#2e2e2e', fg='white', font= ("Verdana", "10"), pady=(50)).pack()

    newWindow.after(1, sendData)

def copyright():
    x = Toplevel(root)
    x.grab_set() 
    x.title("License")
    wx = root.winfo_x() + 100
    wy = root.winfo_y() + 400 - 320
    x.geometry(f"320x300+{wx}+{wy}")
    x.configure(background='#2e2e2e')
    x.iconbitmap(iconfile)
    x.resizable(False, False)
    text=Text(x, padx=10, pady=10)
    text.pack()
    text.config(state="normal")
    text.insert(END, 'Github > github.com/patrick-mns\nCopyright (c) 2022 patrick-mns\n\nRedistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:\n\n1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.\n\n2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.\n\n3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission. \n\nTHIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.')
    text.config(state="disabled")
    return(x)

iconfile = resource_path('src/icon.ico')
root = tk.Tk()
root.resizable(False, False)
root.geometry('500x270+10+20')
root.title("PS Excel - beta")
root.configure(background='#2e2e2e')
root.iconbitmap(iconfile)
menubar = Menu(root)
menubar.add_cascade(label="Autor", command=lambda:[webbrowser.open_new_tab('https://github.com/patrick-mns')])
menubar.add_cascade(label="License", command=lambda:[copyright()])

root.config(menu=menubar)
Label(root, text ="", background='#2e2e2e', fg='white', font= ("Verdana", "15"), pady=(2)).pack()
button_frame = tk.Frame(root, border=4,  bg='#2e2e2e')
button_frame.pack()
labelTextJson = StringVar()
Button(button_frame, text ="Import .xlsx", width=12, bg='#3e3e3e', fg='white', border=0, font= ("Verdana", "10"), pady=10, padx=20, cursor='plus', command= lambda: [select_file('.xlsx', 'select file Excel', labelTextJson)]).pack(side='left')
Label(button_frame, textvariable=labelTextJson, bg='#1e1e1e', fg='white', border=3, font= ("Verdana", "10"), pady=10, padx=20, width=25, justify='left').pack(side='left')
button_frame_psd = tk.Frame(root, border=4,  bg='#2e2e2e')
button_frame_psd.pack()
labelTextPSD = StringVar()
Button(button_frame_psd , text ="Import .psd", width=12, bg='#3e3e3e', fg='white', border=0, font= ("Verdana", "10"), pady=10, padx=20, cursor='plus', command= lambda: [select_file('.psd', 'select file psd', labelTextPSD)]).pack(side='left')
Label(button_frame_psd , textvariable= labelTextPSD, bg='#1e1e1e', fg='white', border=3, font= ("Verdana", "10"), pady=10, padx=20, width=25, justify='left').pack(side='left')
button_frame_psd = tk.Frame(root, border=4,  bg='#2e2e2e')
button_frame_psd.pack() 
labelTextDir = StringVar()
Button(button_frame_psd , text ="Select dir", width=12, bg='#3e3e3e', fg='white', border=0, font= ("Verdana", "10"), pady=10, padx=20, cursor='plus', command= lambda: [select_dir('select a dir', labelTextDir)]).pack(side='left')
Label(button_frame_psd , textvariable= labelTextDir, bg='#1e1e1e', fg='white', border=3, font= ("Verdana", "10"), pady=10, padx=20, width=25, justify='left').pack(side='left')
button_start = tk.Frame(root, border=10,  bg='#2e2e2e')
button_start.pack()
Button(button_start , text ="start", width=46 , bg='#3e3e3e', fg='white', border=0, font= ("Verdana", "10"), pady=10, padx=10, command=data).pack()
root.mainloop()
