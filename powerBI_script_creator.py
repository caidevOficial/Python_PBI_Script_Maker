# Copyright (C) 2023 <FacuFalcone - CaidevOficial>
# 
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
# 
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
# 
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.

import json
import os
from tkinter import *
from tkinter.messagebox import showinfo as alert

import pandas as pd
import customtkinter
import pygame.mixer as mixer 

class PbiScriptMaker(customtkinter.CTk):
    """
    The PbiScriptMaker class creates a Power BI script by replacing placeholders in a template with user
    input and data from a dataframe, and then creates a text file with the script.
    """
    __dataframe = pd.DataFrame()
    __file_paths = {
        "configs": './data/datasets.json',
        "source": './source_pbi_fields.xlsx',
        "destiny": './',
        "sound_error": "./assets/sound/error.mp3",
        "sound_success": "./assets/sound/success.mp3"
    }
    __table_name: str = None
    __dataset_name: str = None
    __full_path: str = None
    
    
    def __init__(self):
        """
        This function initializes a GUI window with various frames, input fields, and buttons for
        creating a Power BI script.
        """
        super().__init__()

        self.title("Power BI Script Creator")
        # Main Frame
        self.__frame_main = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.__frame_main.grid(row=0, column = 0, padx=20, pady=5, columnspan=4, rowspan = 4, sticky="we")
        
        # Banner Img (inside main frame)
        self.__banner = PhotoImage(file='./assets/img/banner_2.png')
        self.__top_banner = Label(master=self.__frame_main, image=self.__banner, text='Banner')
        self.__top_banner.grid_configure(row=0, column=0, padx=20, pady=5, columnspan=5, rowspan=1, sticky='we')
        
        # Secondary Frame (inside main frame) for TextBox, ComboBox & Button
        self.__frame_input = customtkinter.CTkFrame(self.__frame_main, corner_radius=0, fg_color="transparent")
        self.__frame_input.grid(row=1, column = 0, padx=20, pady=5, columnspan=4, rowspan = 2, sticky="nsew")
        
        # Text Box (inside secondary frame)
        self.__txt_table_name = customtkinter.CTkEntry(master=self.__frame_input, height = 14, placeholder_text="Table Name")
        self.__txt_table_name.grid(row=0, column=0, columnspan=3, padx=20, pady=5, sticky="nsew")

        # Combo Box (inside secondary frame)
        self.__combobox_dataset_name = customtkinter.CTkComboBox(master=self.__frame_input, height = 14, width=199, values=self.__open_configs())
        self.__combobox_dataset_name.grid(row=0, column=3, columnspan=2, padx=10, pady=(5, 5), sticky="nsew")

        # Button (inside secondary frame)
        self.__btn_add = customtkinter.CTkButton(master=self.__frame_input, height = 20, text="Create Script", command=self.bttn_create_on_click)
        self.__btn_add.grid(row=1, padx=20, pady=5, columnspan=3, sticky="news")
    
    def __open_configs(self) -> list[str]:
        """
        This function opens a JSON file containing dataset configurations and returns a list of dataset
        names.
        :return: A list of strings containing the names of datasets, which are read from a JSON file
        located at the path specified in the `__file_paths` dictionary under the key `'configs'`.
        """
        with open(self.__file_paths['configs'], 'r') as configs:
            return json.load(configs)['datasets']
    
    def __open_file(self) -> bool:
        """
        This function attempts to open an Excel file and return a boolean value indicating whether it
        was successful or not, raising a FileNotFoundError if it fails.
        :return: a boolean value of True if the Excel file specified in the `__file_paths` dictionary is
        successfully read into a pandas dataframe, and raising a `FileNotFoundError` exception if there
        is an error reading the file.
        """
        try:
            self.__dataframe = pd.read_excel(self.__file_paths['source'])
            return True
        except Exception as e:
            raise FileNotFoundError(e.with_traceback(None))

    def __create_script(self) -> bool:
        """
        This function creates a script in Power Query language for renaming columns in a Google BigQuery
        table.
        :return: a boolean value, either True or False.
        """
        template = """\
    let
        Source = GoogleBigQuery.Database([BillingProject = ProjectID, UseStorageApi = false]),
        Navigation = Source{[Name = DatalakeID]}[Data],
        #"Navigation 1" = Navigation{[Name = "DATASET_NAME", Kind = "Schema"]}[Data],
        #"Navigation 2" = #"Navigation 1"{[Name = "TABLE_NAME", Kind = "Table"]}[Data],
        #"Renamed columns" = Table.RenameColumns(#"Navigation 2", {})
    in
        #"Renamed columns"
    """
        message = ''
        first = True
        schema = '{"TECH", "FUNC"}'
        try:
            self.__table_name = self.__txt_table_name.get()
            self.__dataset_name = self.__combobox_dataset_name.get()
            if not self.__table_name: raise Exception('Empty field Table Name')
            for _, row in self.__dataframe.iterrows():
                replaced = schema.replace('TECH', row['TECH_NAME']).replace('FUNC', row['FUNC_NAME'])
                if first:
                    message += replaced
                    first = False
                else: message += f", {replaced}"
            self.__template = template\
                .replace('(#"Navigation 2", {})', f'(#"Navigation 2", {{{message}}})')\
                .replace('DATASET_NAME', self.__dataset_name)\
                .replace('TABLE_NAME', self.__table_name)
            self.__SEA_messenger('Script Created!', 'info')
            return True
        except Exception as e:
            raise e.with_traceback(None)
    
    def __create_txt_file(self) -> bool:
        """
        This function creates a text file with a given name and writes a template to it, and returns
        True if successful.
        :return: a boolean value, either True or False. In this case, it returns True if the file was
        successfully created and False if an exception was raised.
        """
        try:
            self.__full_path = f"{self.__file_paths['destiny']}{self.__dataset_name}.{self.__table_name}.vba"
            with open(self.__full_path, 'w') as file:
                file.writelines(self.__template)
                self.__SEA_messenger(f"File created: {self.__full_path}", 'success')
                self.__play_sound('sound_success')
                alert("Success", f"File created: {self.__full_path}")
                return True
        except Exception as e:
            raise e.with_traceback(None)
    
    def __SEA_messenger(self, message: str, message_type: str) -> None:
        """
        This is a Python function that prints messages with different colors and message types (error,
        success, information).
        
        :param message: A string containing the message to be displayed
        :param message_type: A string indicating the type of message being passed (e.g. "Error",
        "Success", "Info")
        """
        _b_red: str = '\033[41m'
        _b_green: str = '\033[42m'
        _b_blue: str = '\033[44m'
        _f_white: str = '\033[37m'
        _no_color: str = '\033[0m'
        message_type = message_type.strip().capitalize()
        match message_type:
            case 'Error':
                print(f'{_b_red}{_f_white}> Error: {message}{_no_color}')
            case 'Success':
                print(f'{_b_green}{_f_white}> Success: {message}{_no_color}')
            case 'Info':
                print(f'{_b_blue}{_f_white}> Information: {message}{_no_color}')
    
    def __play_sound(self, audio_name: str) -> None:
        """
        This function plays a sound file specified by the input audio_name using the mixer module in
        Python.
        
        :param audio_name: audio_name is a string parameter that represents the name of the audio file
        to be played. It is used to access the file path of the audio file from a dictionary of file
        paths
        """
        sound = mixer.Sound(self.__file_paths[audio_name])
        mixer.Sound.play(sound)

    def __clear_console(self) -> None:
        """
        This function clears the console screen in Python.
        """
        if os.name in ['ce', 'nt', 'dos']: os.system("cls")
        else: os.system("clear")
        
    def bttn_create_on_click(self) -> None:
        """
        This function initializes a mixer and tries to open a file, create a script, and create a text
        file, catching any exceptions and displaying an error message if necessary.
        """
        self.__clear_console()
        mixer.init()
        
        try:
            self.__open_file()
            self.__create_script()
            self.__create_txt_file()
        except Exception as e:
            self.__play_sound('sound_error')
            self.__SEA_messenger(f'Exception catched: {e}', 'error')
            alert("Error -> Exception", f"Please check the console message!")
        

if __name__ == "__main__":
    PBI_Script_Creator_app = PbiScriptMaker()
    PBI_Script_Creator_app.mainloop()