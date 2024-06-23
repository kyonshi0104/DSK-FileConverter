#GUI
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox as me
from tkinter import filedialog as fid
from tkinter.messagebox import WARNING
#converter
import converter
#outer
import json
from PIL import Image
import threading
import subprocess
from sys import platform
import sys,os
import configparser
from ctypes import windll, byref, create_unicode_buffer, create_string_buffer

def resourcePath(filename):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(filename)

config = configparser.ConfigParser()
config.read('settings.ini')

image_format = ["bmp","dib","eps","gif","icns","ico","im","jpeg","msp","pcx","png","ppm","sgi","spider","tga","tiff","webp","xmb","jpg","pdf","palm"]
image_read_format = ["bmp","jpg","dib","eps","gif","icns","ico","im","jpeg","jpeg2000","msp","pcx","png","ppm","sgi","spider","tga","tiff","webp","xmb","blp","cur","dcx","fli","flc","fpx","frex","gbr","gd","imt","iptc","naa","mcidas","mic","mpo","pcd","pixar","psd","wal","wmf","xpm"]
audio_format = ["mp3","m4v","mp4","avi","flv","wav","ogg","oga","ts","qt","wma","asf","mov","m4a","alac","ape","mac","tta","mka","mkv","flac","aiff","aif","aifc","webm","weba"]
audio_read_format = ["mp3","wav","ogg","oga","ts","wma","asf","mov","m4a","alac","ape","mac","tta","mka","mkv","flac","aiff","aif","aifc"]
video_format = ["mp4","m4v","mp3","avi","flv","wav","ogg","oga","ts","qt","wma","asf","mov","m4a","alac","ape","mac","tta","mka","mkv","flac","aiff","aif","aifc","webm","weba","dat","flv","wmv","asf","mpeg","mpg","vob","mkv","asf","wmv","rm","rmvb","vob"]
video_read_format = ["mp4","m4v","mov","qt","avi","flv","wmv","asf","mpeg","mpg","vob","mkv","asf","wmv","rm","rmvb","vob","ts","dat"]
file_path_i = []
file_path_a = []
file_path_v = []
file_path_str = ""
global_progress_window = None

convert_icon = ctk.CTkImage(Image.open(resourcePath(r"resource\convert.png")),size=(40,40))
clear_icon = ctk.CTkImage(Image.open(resourcePath(r"resource\dustbox.png")),size=(20,20))
allclear_icon = ctk.CTkImage(Image.open(resourcePath(r"resource\dustbox.png")),size=(40,40))

ctk.set_appearance_mode(config.get('theme','mode'))
ctk.set_default_color_theme("dark-blue")

FR_PRIVATE  = 0x10
FR_NOT_ENUM = 0x20

def loadfont(fontpath, private = True, enumerable = False):

    if isinstance(fontpath, str):
        pathbuf = create_string_buffer(fontpath.encode('utf-8'))
        AddFontResourceEx = windll.gdi32.AddFontResourceExA
    else:
        raise TypeError('fontpath must be a str')

    flags = (FR_PRIVATE if private else 0) | (FR_NOT_ENUM if not enumerable else 0)

    numFontsAdded = AddFontResourceEx(byref(pathbuf), flags, 0)

    return bool(numFontsAdded)

loadfont(resourcePath('resource\\PixelMplus12.ttf'))

FONT_TYPE = "PixelMplus12"

class app(ctk.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.geometry("720x500")
        self.title("DSK File Converter / v.1.3.1")
        self.iconbitmap(resourcePath("resource\\icon.ico"))
        self.resizable(False,False)
        self.yotei = ctk.CTkLabel(self,text="DSK File Converter",font=(FONT_TYPE,30))
        self.yotei.grid(row=0,column=0,padx=5,pady=(10,5))
        self.imageframe = image_frame(master=self)
        self.imageframe.grid(row=1,column=0,padx=5,pady=(10,5))
        self.audioframe = audio_frame(master=self)
        self.audioframe.grid(row=2,column=0,padx=5,pady=5)
        self.videoframe = video_frame(master=self)
        self.videoframe.grid(row=3,column=0,padx=5,pady=5)

        self.allclear_button = ctk.CTkButton(self,image=allclear_icon,text="クリア",command=lambda: all_clear(), fg_color="#b22222",hover_color="#800000", corner_radius=10, width=35, height=55,font=(FONT_TYPE,30))
        self.convert_button = ctk.CTkButton(self,image=convert_icon, text="変換",command=lambda: convert(), fg_color="#778899",hover_color="#a9a9a9", corner_radius=10, width=35, height=55,font=(FONT_TYPE,30))
        self.settings_button = ctk.CTkButton(self, text="設定",command=lambda: settings(self), fg_color="#778899",hover_color="#a9a9a9", corner_radius=10, width=35, height=27,font=(FONT_TYPE,15))
        self.settings_button.grid(row=4,column=0,padx=10,pady=(2,5),sticky="sw")
        self.convert_button.grid(row=4,column=0,padx=10,pady=(2,5),sticky="es")
        self.allclear_button.grid(row=4,column=0,padx=(10,150),pady=(2,5),sticky="es")

class image_frame(ctk.CTkFrame):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.imagelabel = ctk.CTkLabel(self,text="画像フォーマット変換",font=(FONT_TYPE,15))
        self.filepath = ctk.CTkTextbox(self,state="disabled",width=550,height=70,font=(FONT_TYPE,15))
        self.dropdown = ctk.CTkComboBox(self,values=image_format,state="readonly")
        self.imagelabel.grid(row=0,column=0,padx=(15,5),pady=5,sticky="w")
        self.filedialog_button = ctk.CTkButton(self,text="ファイルを開く",font=(FONT_TYPE,15))
        self.filedialog_button.configure(command=lambda: file_path_ask(self.filepath,"image"))
        self.imageclear = ctk.CTkButton(self,image=clear_icon,text="クリア",fg_color="#6495ed",hover_color="#800000",font=(FONT_TYPE,15))
        self.imageclear.configure(command=lambda: clear('i'))
        self.filedialog_button.grid(row=0,column=1,padx=5,pady=5)
        self.filepath.grid(row=1,column=0,padx=5,pady=5)
        self.imageclear.grid(row=1,column=1,padx=5,pady=(7,0),sticky="n")
        self.dropdown.grid(row=1,column=1,padx=5,pady=5,sticky="s")

class audio_frame(ctk.CTkFrame):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.audiolabel = ctk.CTkLabel(self,text="音楽フォーマット変換",font=(FONT_TYPE,15))
        self.filepath = ctk.CTkTextbox(self,state="disabled",width=550,height=70,font=(FONT_TYPE,15))
        self.dropdown = ctk.CTkComboBox(self,values=audio_format,state="readonly")
        self.audiolabel.grid(row=0,column=0,padx=(15,5),pady=5,sticky="w")
        self.filedialog_button = ctk.CTkButton(self,text="ファイルを開く",font=(FONT_TYPE,15))
        self.filedialog_button.configure(command=lambda: file_path_ask(self.filepath,"audio"))
        self.audioclear = ctk.CTkButton(self,image=clear_icon,text="クリア",fg_color="#6495ed",hover_color="#800000",font=(FONT_TYPE,15))
        self.audioclear.configure(command=lambda: clear('a'))
        self.filedialog_button.grid(row=0,column=1,padx=5,pady=5)
        self.filepath.grid(row=1,column=0,padx=5,pady=5)
        self.audioclear.grid(row=1,column=1,padx=5,pady=(7,0),sticky="n")
        self.dropdown.grid(row=1,column=1,padx=5,pady=5,sticky="s")

class video_frame(ctk.CTkFrame):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.videolabel = ctk.CTkLabel(self,text="動画フォーマット変換",font=(FONT_TYPE,15))
        self.filepath = ctk.CTkTextbox(self,state="disabled",width=550,height=70,font=(FONT_TYPE,15))
        self.dropdown = ctk.CTkComboBox(self,values=video_format,state="readonly")
        self.videolabel.grid(row=0,column=0,padx=(15,5),pady=5,sticky="w")
        self.filedialog_button = ctk.CTkButton(self,text="ファイルを開く",font=(FONT_TYPE,15))
        self.filedialog_button.configure(command=lambda: file_path_ask(self.filepath,"video"))
        self.videoclear = ctk.CTkButton(self,image=clear_icon,text="クリア",fg_color="#6495ed",hover_color="#800000",font=(FONT_TYPE,15))
        self.videoclear.configure(command=lambda: clear('v'))
        self.filedialog_button.grid(row=0,column=1,padx=5,pady=5)
        self.filepath.grid(row=1,column=0,padx=5,pady=5)
        self.videoclear.grid(row=1,column=1,padx=5,pady=(7,0),sticky="n")
        self.dropdown.grid(row=1,column=1,padx=5,pady=5,sticky="s")

#progress

class progress(ctk.CTkToplevel):
    def __init__(self, master, *args, **kwargs):
        super().__init__(*args, **kwargs)
        popup_window_geo = (f"300x50+{str(self.master.winfo_x()+210)}+{str(self.master.winfo_y()+250)}")
        self.geometry(popup_window_geo)
        self.wm_overrideredirect(True)
        self.iconbitmap(resourcePath("resource\\icon.ico"))
        self.resizable(False,False)
        self.grab_set()

        self.progress = ctk.CTkProgressBar(self,width=270,height=5)
        self.progress_label = ctk.CTkLabel(self,text="ファイルを変換中....  0/0",font=(FONT_TYPE,15))
        self.progress_label.grid(row=0,padx=10,pady=(10,0))
        self.progress.grid(row=1,padx=10,pady=10,sticky="n")

class settings(ctk.CTkToplevel):
    def __init__(self, master, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.geometry("350x310")
        self.title("DSK File Converter / 設定")
        self.wm_iconbitmap(resourcePath("resource\\icon.ico"))
        self.resizable(False,False)
        self.grab_set()

        if platform.startswith("win"):
            self.after(200, lambda: self.iconbitmap(resourcePath("resource\\icon.ico")))

        self.f = settings_f(master=self)
        self.t = settings_t(master=self)
        self.o = settings_o(master=self)
        self.f.grid(row=0,column=0,padx=5,pady=5)
        self.t.grid(row=1,column=0,padx=5,pady=5)
        self.o.grid(row=2,column=0,padx=5,pady=5,sticky="news")

#settings frame

class settings_f(ctk.CTkFrame):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        print(os.path.abspath("save").replace("\\","/"))
        self.d_wlavel = ctk.CTkLabel(self,text="保存するフォルダ",font=(FONT_TYPE,15))
        self.d_box = ctk.CTkTextbox(self,font=(FONT_TYPE,15),width=325,height=20)
        if config.get("path","output") == os.path.abspath("save").replace("\\","/") or config.get("path","output") == "save":
            self.d_box.insert('0.0','既定')
        else:
            self.d_box.insert('0.0',extract_and_ellipsis(config.get("path","output")))
        self.d_box.configure(state="disabled")
        self.h_butt = ctk.CTkButton(self,text="変更",font=(FONT_TYPE,15),width=50,height=20,bg_color="#333333",command=lambda: save_filepath_change(self.d_box))
        self.d_wlavel.grid(row=0,column=0,padx=(20,10),pady=5,sticky="w")
        self.d_box.grid(row=1,column=0,columnspan=2,padx=(10,5),pady=5)
        self.h_butt.grid(row=1,column=1,padx=(10,10),pady=5,sticky="e")

class settings_t(ctk.CTkFrame):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.d_wlavel = ctk.CTkLabel(self,text="テーマ",font=(FONT_TYPE,15))
        self.d_box = ctk.CTkComboBox(self,font=(FONT_TYPE,15),width=325,height=30,values=["dark","light"],state="readonly")
        self.d_box.set(config.get("theme","mode"))
        self.d_box.configure(command=lambda  choice=self.d_box.get(): theme_change(choice))
        self.d_wlavel.grid(row=0,column=0,padx=(20,10),pady=5,sticky="w")
        self.d_box.grid(row=1,column=0,columnspan=2,padx=(10,5),pady=5)

class settings_o(ctk.CTkFrame):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.d_wlavel = ctk.CTkLabel(self,text="その他",font=(FONT_TYPE,13))
        self.ca = ctk.CTkCheckBox(self,text='一括クリアボタンを押した際に警告を表示する',font=(FONT_TYPE,12),corner_radius=5, onvalue="on", offvalue="off")
        self.el = ctk.CTkCheckBox(self,text='変換後にエクスプローラを開く',font=(FONT_TYPE,12),corner_radius=5, onvalue="on", offvalue="off")
        self.ed = ctk.CTkCheckBox(self,text='エラーの詳細を表示する',font=(FONT_TYPE,12),corner_radius=5, onvalue="on", offvalue="off")
        self.ca.configure(command=lambda setting = "ca",choice = self.ca: setting_change(setting,choice))
        self.el.configure(command=lambda setting = "el",choice = self.el: setting_change(setting,choice))
        self.ed.configure(command=lambda setting = "ed",choice = self.ed: setting_change(setting,choice))
        if config.get('settings','clear_attention') == "True":
            self.ca.select()
        if config.get('settings','explorer_launch') == "True":
            self.el.select()
        if config.get('settings','error_detail') == "True":
            self.ed.select()
        self.d_wlavel.grid(row=0,column=0,padx=(20,10),pady=(1,1),sticky="nw")
        self.ca.grid(row=1,column=0,columnspan=2,padx=(5,5),pady=(0,5),sticky="nw")
        self.el.grid(row=2,column=0,columnspan=2,padx=(5,5),pady=(0,5),sticky="nw")
        self.ed.grid(row=3,column=0,columnspan=2,padx=(5,5),pady=(0,5),sticky="nw")

#processing

def setting_change(setting,choice:ctk.CTkCheckBox):
    print(choice.get())
    result = 'True' if choice.get() == "on" else 'False'
    if setting == "ca":
        config["settings"]["clear_attention"] = result
        print(f'ca change {result}')
    if setting == "el":
        config["settings"]["explorer_launch"] = result
        print(f'el change {result}')
    if setting == "ed":
        config["settings"]["error_detail"] = result
        print(f'ed change {result}')
    with open('settings.ini', 'w') as configfile:
        config.write(configfile)
        print(f'config write')

def theme_change(choice):
    print('theme change suru dosue')
    ctk.set_appearance_mode(choice)
    config["theme"]["mode"] = choice
    with open('settings.ini', 'w') as configfile:
        config.write(configfile)

def save_filepath_change(textbox:ctk.CTkTextbox):
    d = fid.askdirectory()
    if not d:
        return
    config["path"]["output"] = d
    with open('settings.ini', 'w') as configfile:
        config.write(configfile)
    textbox.configure(state="normal")
    textbox.delete("1.0","end")
    if config.get("path","output") == os.path.abspath("save").replace("\\","/") or config.get("path","output") == "save":
            textbox.insert('0.0','既定')
    else:
        textbox.insert('0.0',extract_and_ellipsis(d))
    textbox.configure(state="disabled")

def file_path_ask(textbox: ctk.CTkTextbox, mode : str):
    global file_path_i
    global file_path_a
    global file_path_v
    if mode == "image":
        if not file_path_i:
            file_path_i = fid.askopenfilenames(filetypes=[("Image file", " .".join(image_read_format).removeprefix(" "))])
        else:
            file_path_i += fid.askopenfilenames(filetypes=[("Image file", " .".join(image_read_format).removeprefix(" "))])
        file_path_str = " , ".join([name.split("/")[-1] for name in file_path_i])
    elif mode == "audio":
        if not file_path_a:
            file_path_a = fid.askopenfilenames(filetypes=[("Audio file", " .".join(audio_read_format).removeprefix(" "))])
        else:
            file_path_a += fid.askopenfilenames(filetypes=[("Audio file", " .".join(audio_read_format).removeprefix(" "))])
        file_path_str = " , ".join([name.split("/")[-1] for name in file_path_a])
    elif mode == "video":
        if not file_path_v:
            file_path_v = fid.askopenfilenames(filetypes=[("Audio file", " .".join(video_read_format).removeprefix(" "))])
        else:
            file_path_v += fid.askopenfilenames(filetypes=[("Audio file", " .".join(video_read_format).removeprefix(" "))])
        file_path_str = " , ".join([name.split("/")[-1] for name in file_path_v])
    else:
        return False
    textbox.configure(state="normal")
    textbox.delete("1.0","end")
    textbox.insert("0.0", file_path_str)
    textbox.configure(state="disabled")
    return True

def convert():
        setting = settings_correctly()
        if setting == "FileNotFound":
            me.showwarning("エラー","ファイルが一つも選択されていません。")
        elif setting == "dropdown not correctly":
            me.showwarning("エラー","変換する拡張子の指定がされていません。")
        elif setting == "OutputdirNotFound":
            me.showerror("エラー","出力先のフォルダが見つかりませんでした。\n移動または削除された可能性があります。\n「設定」より、新たな保存先フォルダを指定してください。")
        else:
            con_list = list(file_path_i)+list(file_path_a)+list(file_path_v)
            berw = progress(app)
            berw.progress.set(0)
            berw.progress_label.configure(text=f"ファイルを変換中....  0/{str(len(con_list))}")
            output_path = config.get('path','output')
            convert_thread = threading.Thread(target=file_converts,args=(berw.progress,berw.progress_label,list(file_path_i),list(file_path_a),list(file_path_v),App.imageframe.dropdown.get(),App.audioframe.dropdown.get(),App.videoframe.dropdown.get(),output_path))
            convert_thread.start()

def extract_and_ellipsis(input_string: str, end_position: int = 25):
    if len(input_string) < end_position+1 :
        return input_string
    else:
        extracted_substring = input_string[:end_position]
        return f"{extracted_substring}..."

def file_converts(ber:ctk.CTkProgressBar,label:ctk.CTkLabel,i_list:list,a_list:list,v_list:list,i_ex:str,a_ex:str,v_ex:str,outout_path:str):
    App.attributes("-disabled",True)
    length = len(i_list+a_list+v_list)
    leng_con = 0
    one_bg = 1 / length
    ber_pro = 0
    error_FE = 0
    error_FN = 0
    error_OT = 0
    completed = 0
    ber.set(0)
    for i in i_list:
        result = converter.image_convert(i,f"{outout_path}\\{i.split("/")[-1].split(".")[0]}.{i_ex}")
        if result == "FileExistsError":
            error_FE += 1
        elif result == "FileNotFoundError":
            error_FN += 1
            print(f'FNdosue\n{f"outout_path\\{i.split("/")[-1].split(".")[0]}.{i_ex}"}\n{result}')
        elif result == False:
            error_OT += 1
            print(f'OTdosue\n{f"outout_path\\{i.split("/")[-1].split(".")[0]}.{i_ex}"}\n{result}')
        else:
            completed += 1
            print('OKdosue')
        leng_con += 1
        ber_pro += one_bg
        ber.set(ber_pro)
        label.configure(text=f'ファイルを変換中....  {str(leng_con)}/{str(length)}')
    for a in a_list:
        print(a)
        result = converter.audio_convert(a,f"{outout_path}\\{a.split("/")[-1].split(".")[0]}.{a_ex}")
        if result == "FileExistsError":
            error_FE += 1
        elif result == "FileNotFoundError":
            error_FN += 1
            print(f'FNdosue\n{f"outout_path\\{a.split("/")[-1].split(".")[0]}.{a_ex}"}\n{result}')
        elif result == False:
            error_OT += 1
            print(f'OTdosue\n{f"outout_path\\{a.split("/")[-1].split(".")[0]}.{a_ex}"}\n{result}')
        else:
            completed += 1
            print('OKdosue')
        leng_con += 1
        ber_pro += one_bg
        ber.set(ber_pro)
        label.configure(text=f'ファイルを変換中....  {str(leng_con)}/{str(length)}')
        result = converter.audio_convert
    for v in v_list:
        print(v)
        result = converter.video_convert(v,f"{outout_path}\\{v.split("/")[-1].split(".")[0]}.{v_ex}")
        if result == "FileExistsError":
            error_FE += 1
        elif result == "FileNotFoundError":
            error_FN += 1
            print(f'FNdosue\n{f"outout_path\\{v.split("/")[-1].split(".")[0]}.{v_ex}"}\n{result}')
        elif result == False:
            error_OT += 1
            print(f'OTdosue\n{f"outout_path\\{v.split("/")[-1].split(".")[0]}.{v_ex}"}\n{result}')
        else:
            completed += 1
            print('OKdosue')
        leng_con += 1
        ber_pro += one_bg
        ber.set(ber_pro)
        label.configure(text=f'ファイルを変換中....  {str(leng_con)}/{str(length)}')
        result = converter.audio_convert
    if config.get('settings','error_detail') == "True":
        me.showinfo("完了",f"全てのファイルのコンバートが完了しました！\nエラー：\nFileExistsError：{error_FE}\nFileNotFoundError：{error_FN}\nOtherError：{error_OT}")
    else:
        me.showinfo("完了",f"全てのファイルのコンバートが完了しました！\nエラー：{error_FE+error_FN+error_OT}")
    App.attributes("-disabled",False)
    ber.master.destroy()
    all_clear(False)
    if config.get('settings','explorer_launch') == "True":
        explorer(outout_path,i_list,a_list,v_list,i_ex,a_ex,v_ex)

def settings_correctly():
    if not list(file_path_i)+list(file_path_a)+list(file_path_v):
        return "FileNotFound"
    elif file_path_i and App.imageframe.dropdown.get() == "" or file_path_a and App.audioframe.dropdown.get() == "" or file_path_v and App.videoframe.dropdown.get() == "":
        return "dropdown not correctly"
    elif os.path.isdir(config.get("path","output")) == False:
        return "OutputdirNotFound"
    else:
        return "Settings all correctly"

def all_clear(d=True):
    if config.get('settings','clear_attention') == "True" and d == True:
        result = me.askokcancel('警告','削除してもよろしいですか？',icon=WARNING)
    else:
        result = True
    if result:
        global file_path_i
        global file_path_a
        global file_path_v
        i = App.imageframe.filepath
        a = App.audioframe.filepath
        v = App.videoframe.filepath
        i.configure(state="normal")
        a.configure(state="normal")
        v.configure(state="normal")
        i.delete("1.0","end")
        a.delete("1.0","end")
        v.delete("1.0","end")
        i.configure(state="disabled")
        a.configure(state="disabled")
        v.configure(state="disabled")
        file_path_i = []
        file_path_a = []
        file_path_v = []

def clear(mode:str):
    if mode == "i":
        global file_path_i
        m = App.imageframe.filepath
        file_path_i = []
    elif mode == "a":
        global file_path_a
        m = App.audioframe.filepath
        file_path_a = []
    elif mode == "v":
        global file_path_v
        m = App.videoframe.filepath
        file_path_v = []
    else:
        return
    m.configure(state="normal")
    m.delete("1.0","end")
    m.configure(state="disabled")

def explorer(output_path:str,i_list:list,a_list:list,v_list:list,i_ex:str,a_ex:str,v_ex:str):
    if i_list:
        subprocess.Popen(f'explorer /select,"{output_path.replace('/','\\')}\\{i_list[0].split("/")[-1].split(".")[0]}.{i_ex}"')
        print(f'explorer /select,"{output_path}/{i_list[0].split("/")[-1].split(".")[0]}.{i_ex}"')
    elif a_list:
        subprocess.Popen(f'explorer /select,"{output_path.replace('/','\\')}\\{a_list[0].split("/")[-1].split(".")[0]}.{a_ex}"')
        print(f'explorer /select,"{output_path}/{a_list[0].split("/")[-1].split(".")[0]}.{a_ex}"')
    elif v_list:
        subprocess.Popen(f'explorer /select,"{output_path.replace('/','\\')}\\{v_list[0].split("/")[-1].split(".")[0]}.{v_ex}"')
        print(f'explorer /select,"{output_path}/{v_list[0].split("/")[-1].split(".")[0]}.{v_ex}"')

App = app()
App.mainloop()