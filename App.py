import tkinter.filedialog
from tkinter import filedialog

import customtkinter as ctk
from Functions import *
import threading
from settings import *


class App(ctk.CTk):
    def __init__(self):
        super().__init__(fg_color='#14171A')
        # Variables
        self.link = ctk.StringVar(self,
                                  value="https://oauth.vk.com/blank.html#access_token=vk1.a.N6k9S9XkHfGvoAxZf6b8NGBRm-6EKftc3GjUsCBgvLe8rjbdbC_nXb4c2Mx18Z8fUvdjKaDqYkyrZkJvlAPV2fIuerW3FToUbOTHctI_dPAtG_yU6--EnNuNRoMu0dbOrQL9V07FX4ML3FWWXIOT2k-a_lMOwvVw8Eg_oNZ64asQ9ptqTA3BG6vc1CwuFiRLN4CufB5AhtmaiUqirGRKzw&expires_in=86400&user_id=533656017")
        self.excel_path = ctk.StringVar(self, value="vkApi_excel_otchyot_Za_May.xlsx")
        self.path_to_save = ctk.StringVar(self, value="vkApi_excel_save.xlsx")
        self.sheet_number = ctk.IntVar(self, value=2)
        self.did_anything_change = ctk.BooleanVar(self, value=False)
        self.chat_option = ctk.StringVar(self, value="Dance chat")
        # Basics
        self.title("Poll app")
        self.size_of_window = f"650x350"
        self.geometry(self.size_of_window)
        self.minsize(height=int(self.size_of_window[4:]),
                     width=int(self.size_of_window[:3]))
        self.maxsize(height=int(self.size_of_window[4:]),
                     width=int(self.size_of_window[:3]))

        # Layout
        self.rowconfigure((0, 1, 2), weight=1, uniform='a')
        self.columnconfigure((0, 1), weight=1, uniform='a')

        # Widgets
        self.logs = Logs(self)
        self.menu = Menu(self)
        # ctk.CTkFrame(self, fg_color="blue", height=120).grid(row=1, column=0, columnspan=2, sticky='nsew')
        # ctk.CTkFrame(self, fg_color="red", height=100).grid(row=2, column=0, columnspan=2, sticky='nsew')
        # self.logs.tkraise()
        # self.menu.tkraise()
        self.mainloop()


class Logs(ctk.CTkTextbox):
    def __init__(self, parent):
        super().__init__(parent, fg_color=TEXTBOX_COLOR_DARK, text_color='white', wrap="none",
                         border_color=TEXTBOX_BORDER_COLOR_DARK, border_width=BORDER_WIDTH)

        # self.excel_path = parent.excel_path
        # self.path_to_save = parent.path_to_save
        # self.sheet_number = parent.sheet_number

        self.grid(row=0, rowspan=3, column=1, sticky='nsew', pady=15, padx=10)


class Menu(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, fg_color='black')
        self.chat_option = None
        self.excel_path_to_open = ctk.StringVar(self, value="vkApi_excel_otchyot_Za_May.xlsx")
        self.sheet_action = False
        self.link = parent.link
        self.grid(row=0, rowspan=3, column=0, sticky='nsew', pady=15, padx=10)
        # Layout
        self.rowconfigure((2, 3, 4, 5, 6), weight=1, uniform='a')
        self.master.rowconfigure(1, weight=0)
        self.master.rowconfigure(0, weight=0)
        self.columnconfigure((0, 1, 2), weight=1, uniform='a')

        # Widgets
        token_widgets(self)
        ctk.CTkLabel(self, text="Choose File to add to").grid(row=3, column=0, columnspan=2, sticky='nw', padx=5)
        ctk.CTkButton(self, corner_radius=CORNER_RADIUS, text="open", command=self.open_filedialog,
                      fg_color=BUTTON_COLOR_DARK, height=35).grid(row=3, column=0, sticky='sew', padx=5)
        ctk.CTkLabel(self, corner_radius=CORNER_RADIUS, textvariable=self.excel_path_to_open, bg_color=LABEL_COLOR_DARK,
                     height=35).grid(
            row=3, column=1,
            columnspan=3,
            sticky='sew',
            padx=5)
        two_buttons(self)

        ctk.CTkCheckBox(self, variable=parent.did_anything_change, text="Did the crew change?").grid(row=5, column=0,
                                                                                                     columnspan=3,
                                                                                                     padx=5,
                                                                                                     sticky='n')
        ctk.CTkOptionMenu(self, values=["Dance chat", "Fence chat"], variable=parent.chat_option).grid(row=5, column=0,
                                                                                              columnspan=3, padx=5,
                                                                                              sticky='s')
        self.gol_button = ctk.CTkButton(master=self,
                      text="ГОООООООЛЛЛЛЛЛЛ!!!!",
                      command=lambda: self.threading_everything(parent),
                      fg_color=BUTTON_COLOR_DARK,
                      border_width=BORDER_WIDTH,
                      state=tkinter.DISABLED,
                      border_color=BUTTON_BORDER_COLOR_DARK)
        self.gol_button.grid(row=6, column=0, columnspan=3, sticky='nsew', pady=10,
                                                                  padx=5)


    def threading_everything(self, parent):
        t1 = threading.Thread(target=insert_all_the_polls, args=(parent.excel_path.get(),
                                                                 parent.path_to_save.get(),
                                                                 parent.link.get(),
                                                                 parent.logs,
                                                                 parent.did_anything_change.get(),
                                                                 parent.chat_option.get(),
                                                                 self.sheet_action))
        t1.start()

    def open_filedialog(self):
        # path = tkinter.filedialog.askopenfile(title="Select a File")
        path = filedialog.askopenfilename(title="Select a File")
        if len(path) == 0:
            return
        self.excel_path_to_open.set(value=path.name[path.name.rfind('/') + 1:])


class two_buttons(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, fg_color='black')
        self.grid(row=4, column=0, columnspan=3, sticky='nsew')
        self.rowconfigure(0, weight=1, uniform='b')
        self.columnconfigure((0, 1), weight=1, uniform='b')
        self.create_new = ctk.CTkButton(self, corner_radius=CORNER_RADIUS, text="Create a new sheet",
                                        command=lambda: self.button_pressed(parent, False), fg_color=BUTTON_COLOR_DARK,
                                        border_color=BUTTON_BORDER_COLOR_DARK, border_width=BORDER_WIDTH)
        self.add_to_the_old = ctk.CTkButton(self, corner_radius=CORNER_RADIUS, text="Add to the last sheet",
                                            command=lambda: self.button_pressed(parent, True), fg_color=BUTTON_COLOR_DARK,
                                            border_color=BUTTON_BORDER_COLOR_DARK, border_width=BORDER_WIDTH)

        self.add_to_the_old.grid(row=0, column=1, sticky='nsew', padx=5, pady=10)
        self.create_new.grid(row=0, column=0, sticky='nsew', padx=5, pady=10)

    def button_create_new_pressed(self):
        self.create_new.configure(state=tkinter.ACTIVE, fg_color=BUTTON_ACTIVE_DARK)
        self.add_to_the_old.configure(fg_color=BUTTON_DISABLED_DARK)

    def button_add_to_the_old_pressed(self):
        self.add_to_the_old.configure(state=tkinter.ACTIVE, fg_color=BUTTON_ACTIVE_DARK)
        self.create_new.configure(fg_color=BUTTON_DISABLED_DARK)
    def button_pressed(self, parent, status):
        parent.gol_button.configure(state=tkinter.NORMAL)
        if status:
            a = self.create_new
            b = self.add_to_the_old
            parent.sheet_action = False
        else:
            a = self.add_to_the_old
            b = self.create_new
            parent.sheet_action = True
        b.configure(state=tkinter.ACTIVE, fg_color=BUTTON_ACTIVE_DARK)
        a.configure(fg_color=BUTTON_DISABLED_DARK)

class token_widgets(ctk.CTkFrame):
    def __init__(self, parent):
        # Init
        super().__init__(parent, fg_color='black')
        self.grid(row=2, column=0, columnspan=3, sticky='nswe')
        # Layout
        self.rowconfigure(list(range(3)), weight=1, uniform='c')
        self.columnconfigure(list(range(3)), weight=1, uniform='c')
        # Widgets
        ctk.CTkLabel(self, text="Enter vk token").grid(row=0, column=0, columnspan=2, sticky='sw', padx=5)
        ctk.CTkEntry(self,
                     corner_radius=CORNER_RADIUS,
                     bg_color="black",
                     fg_color=TEXTBOX_COLOR_DARK,
                     textvariable=parent.link).grid(row=1,
                                                         rowspan=2,
                                                         column=0,
                                                         columnspan=3,
                                                         sticky='nwe',
                                                         padx=5,
                                                         pady=5)
        # ctk.CTkLabel(self, text="Choose File to add to").grid(row=3, column=0, columnspan=2, sticky='sw', padx=5)

App()
