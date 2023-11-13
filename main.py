import ttkbootstrap as ttk
import tkinter as tk
import pandas as pd
import os
from tkinter import filedialog
from ttkbootstrap.constants import *
from allpairspy import AllPairs
from ttkbootstrap.dialogs import Messagebox
from pathlib import Path
from itertools import cycle
from PIL import Image, ImageTk, ImageSequence
from ttkbootstrap import Style
from threading import Thread
from time import sleep


class MainScreen(ttk.Frame):
    def __init__(self, root, switch_frame, *args, **kwargs):
        ttk.Frame.__init__(self, root, *args, **kwargs)
        self.pack(fill="both", expand=True)
        self.excel_path = tk.StringVar()
        self.result_path = tk.StringVar()
        self.switch_frame = switch_frame
        self.create_widgets()

    def create_widgets(self):
        # Labels
        self.title = ttk.Label(self, text="All Pairs!", font=('Georgia', 50))
        self.label_excel = ttk.Label(self, text="EXCEL PATH")
        self.label_result_path = ttk.Label(self, text="RESULT PATH")
        self.label_open_when_finished = ttk.Label(self, text="Open results folder when finished")

        # Entrys
        self.entry_excel_path = ttk.Entry(self, state="disabled", textvariable=self.excel_path)
        self.entry_result_path = ttk.Entry(self, state="disabled", textvariable=self.result_path)

        # CheckButtons
        self.checkbutton_open_when_finished = ttk.Checkbutton(self, bootstyle="square-toggle")
        
        # Buttons
        self.button_start = ttk.Button(self, text="GET ALL PAIRS",style='Outline.TButton', command=self.on_start_button_clicked)
        self.button_select_excel = ttk.Button(self, text="Select excel", style='Outline.TButton', command=self.select_excel_file)
        self.button_select_result_folder = ttk.Button(self, text="Select results folder", style='Outline.TButton', command=self.select_result_folder)

        # Display
        self.title.pack(pady=(20,0))
        self.label_excel.pack(fill="both", padx=30)
        self.entry_excel_path.pack(fill="both", padx=30)
        self.button_select_excel.pack(fill="both", padx=30, pady=(1,0))
        self.label_result_path.pack(fill="both", padx=30, pady=(10,0))
        self.entry_result_path.pack(fill="both", padx=30)
        self.button_select_result_folder.pack(fill="both", padx=30, pady=(1,0))
        self.label_open_when_finished.pack(fill="both", padx=30, pady=(10,0))
        self.checkbutton_open_when_finished.pack(fill="both", padx=30)
        self.button_start.pack(pady=(10,20))

    def select_excel_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel File", "*.xlsx;*.xls")]) 
        self.excel_path.set(file_path)
        self.excel = pd.ExcelFile(self.excel_path.get())
        self.excel_name = self.excel_path.get().split("/")[-1].split(".")[0]

    def select_result_folder(self):
        file_path = filedialog.askdirectory()
        self.result_path.set(file_path)

    def process_excel_file(self):
        # Sleep 1 second to make time for the loading screen to appear
        sleep(1)

        self.button_start.config(state="disabled")
        self.button_select_excel.config(state="disabled")
        self.button_select_result_folder.config(state="disabled")
        self.checkbutton_open_when_finished.config(state="disabled")

        df= pd.read_excel(self.excel)

        # Make list of columns instead of rows and drop nan values
        column_lists = [df[column].tolist() for column in df.columns]
        cleaned_list = [[value for value in row if not pd.isna(value)] for row in column_lists]

        # Get all pairs and export to result path
        result = [pairs for pairs in AllPairs(cleaned_list)]
        df_result = pd.DataFrame(result, columns=[df.columns])
        result_file_path = f'{self.result_path.get()}/{self.excel_name}_allpairs.csv'
        df_result.to_csv(result_file_path, index=False)        

    def on_start_button_clicked(self):
        if self.excel_path.get() == "":
            mb = Messagebox.ok(title="Alert", message="You didn't select an Excel file")
            raise ValueError(title="Alert", message="You didn't select an Excel file")
            
        if self.result_path.get() == "":
            mb = Messagebox.ok("You didn't select a result folder")
            raise ValueError("You didn't select a result folder")
        
        # Switch to loading screen
        self.switch_frame(LoadingScreen)

        # Start new thread to prevent UI from freezing
        new_thread = Thread(target=self.process_excel_file,
                            daemon=True)
        new_thread.start()

        # Check thread status periodically without blocking the UI
        self.check_thread_status(new_thread)

    def check_thread_status(self, thread):
        if thread.is_alive():
            # If the thread is still running, check again after a short delay
            self.after(100, lambda: self.check_thread_status(thread))
        else:
            self.button_start.config(state="enabled")
            self.button_select_excel.config(state="enabled")
            self.button_select_result_folder.config(state="enabled")
            self.checkbutton_open_when_finished.config(state="enabled")

            self.switch_frame(MainScreen)

            # Display success message
            mb = Messagebox.ok(title="Success!", message="All Pairs were generated successfully!")
            print(mb)

            # Open file folder if needed
            if self.checkbutton_open_when_finished.state()[0] == 'selected':
                path = os.path.realpath(self.result_path.get())
                os.startfile(path)


class LoadingScreen(ttk.Frame):
    def __init__(self, root, switch_frame, *args, **kwargs):
        ttk.Frame.__init__(self, root, *args, **kwargs)
        self.pack(fill="both", expand=True)
        self.switch_frame = switch_frame
        self.create_widgets()
    
    def create_widgets(self):
        # open the GIF and create a cycle iterator
        file_path = Path(__file__).parent / "assets/spinners.gif"
        with Image.open(file_path) as im:
            # create a sequence
            sequence = ImageSequence.Iterator(im)
            images = [ImageTk.PhotoImage(s) for s in sequence]
            self.image_cycle = cycle(images)

            # length of each frame
            self.framerate = im.info["duration"]

        self.img_container = ttk.Label(self, image=next(self.image_cycle))
        self.img_container.pack(fill="both", expand="yes")
        self.after(self.framerate, self.next_frame)

    def next_frame(self):
        """Update the image for each frame"""
        self.img_container.configure(image=next(self.image_cycle))
        self.after(self.framerate, self.next_frame)


class Application(ttk.Window):
    def __init__(self):
        super().__init__()
        self.geometry("500x400")
        self.resizable(False, False)
        self.title("All Pairs UI")
        self.themename = Style('solar')  
        self.frames = {}
        self.create_frames()
        self.show_frame(MainScreen)
        self.mainloop()

    def create_frames(self):
        for FrameClass in [MainScreen, LoadingScreen]:
            frame = FrameClass(self, self.switch_frame)
            self.frames[FrameClass] = frame

    def show_frame(self, container):
        for frame in self.frames.values():
            frame.pack_forget()
        self.frames[container].pack(fill="both", expand=True)

    def switch_frame(self, frame_class):
        new_frame = self.frames[frame_class]
        new_frame.update() if hasattr(new_frame, 'update') else None
        self.show_frame(frame_class)

if __name__ == "__main__":
    app = Application()
