import logging
import tempfile
import threading
from multiprocessing import Process
import pywinauto
from pywinauto import Application
import win32com.client as win32
from comtypes.client import GetActiveObject, CreateObject
import concurrent.futures
import pythoncom
import gc
import time
import datetime
from pathlib import Path
from bmk.functions import list_args, gen_tempfile_name
import pandas as pd
import numpy as np
import threading
import time


QLIKVIEW_PROTOCOL: str = 'qvp://'
BMK_SERVER_URL: str = QLIKVIEW_PROTOCOL + '172.16.2.20'
GIPPO_SERVER_URL: str = QLIKVIEW_PROTOCOL + 'qlikview'
QLIKVIEW_SERVER_LOGIN: str  = 'kdad'
QLIKVIEW_SERVER_PASSWORD: str  = 'As1e29fo5_1c'
QLIKVIEW_PATH: str = "C:\Program Files\QlikView\Qv.exe"
QLIKVIEW_USER: str = 'kaaa'
QLIKVIEW_USER_PASSWORD: str = 'user'

class App:
    """A class representing a running instance of a QlikView app."""

    def __init__(self) -> None:
        """Initialize an instance."""
        try:  # if QlikView is open try comtypes.client.GetActiveObject() to connect
            self.com: object = GetActiveObject('QlikTech.QlikView')
        # Qlik is open, but comtypes.client.GetActiveObject() failed - use win32 to connect
        except AttributeError:  
            self.com: object = win32.DispatchEx("QlikTech.QlikView")       
        except OSError:  # Qlik is closed
            try:  # try comtypes.client.CreateObject() to open the app
                self.com: object = CreateObject('QlikTech.QlikView')
            except AttributeError:  # comtypes.client.CreateObject() failed - use win32 to open
                self.com: object = win32.DispatchEx("QlikTech.QlikView")
        


        self.uia: object = Application(backend='uia').connect(path=QLIKVIEW_PATH) # specify keyword directly (write path=) for connect to work
        self.pid: int = self.com.GetProcessId()
        self.servers = (BMK_SERVER_URL, GIPPO_SERVER_URL)
    
    def list_docs(self) -> None:
        """List all docs on the specified servers."""
        for server in self.servers:
            server_doc_list = self.com.GetServerDocList(server) # IArrayOfDocListEntry
            print(f"SERVER: '{server}'")
            print("DOCS ON SERVER:")
            for i in range(server_doc_list.Count):
                doc_name = server_doc_list.Item(i).DocName
                print(f"'{server}/{doc_name}'")
            print()
        
    def open_doc(self, doc_path: str) -> object:
        """Open a QlikView document and return a Document object, then create a QlikDoc object based on it."""

        doc_login_path = f"{QLIKVIEW_PROTOCOL}{QLIKVIEW_SERVER_LOGIN}@{doc_path.replace(QLIKVIEW_PROTOCOL, '')}"

        
        print(f"Doc login path is {doc_login_path}")
        def call_open_doc(doc_login_path: str) -> object:
            """Call an OpenDoc method from the app's COM object."""
            return self.com.OpenDoc(doc_login_path, QLIKVIEW_USER, QLIKVIEW_USER_PASSWORD)
         

        def login_to_doc() -> None:
            """Perform a sequence of GUI actions for logging in to the document."""
            if  self.uia.top_window()[f'Password'].exists(): # server login is not necessary if we have already preformed a connection to the server during this session
                self.uia.top_window()[f'Password'].Edit.type_keys(QLIKVIEW_SERVER_PASSWORD)
                self.uia.top_window()['OK'].click()
            
            self.uia.top_window()['User Identification'].type_keys(QLIKVIEW_USER)
            self.uia.top_window()['OK'].click()
            self.uia.top_window()['Password'].type_keys(QLIKVIEW_USER_PASSWORD)
            self.uia.top_window()['OK'].click()

   
        
        # we have to use multi threading in order to be able to use GUI automation while calling OpenDoc()
        with concurrent.futures.ThreadPoolExecutor() as ex:
            ex.submit(login_to_doc)
            result = call_open_doc(doc_login_path) # does not work as part of the pool
            
        return Doc(doc_path, result)

              
class Doc():
    """A class representing a QlikView document."""
 
    def __init__(self, doc_path: str, com: object) -> object:
        """Initialize the class"""
        self.path: str = doc_path
        self.com: object = com # @ https://stackoverflow.com/questions/625083/what-init-and-self-do-in-python
        self.field_count: int = self.com.GetFieldCount()
        self.com.Clear()

    def __repr__(self) -> str:
        """Give some info about an object instance."""
        return f"<qlik.Doc - '{self.path}', field_count={self.field_count}>"

    def list_sheets(self) -> list['str']:
        """List all sheets' names in the document."""
        for i in range(self.com.NoOfSheets()):
            print(f"{self.com.Sheets(str(i)).GetProperties().Name!r}")
    
    def get_sheet(self, sheet_name: str) -> object:
        """Get a sheet COM object by *name* and create a Sheet object from it."""
        return Sheet(self.com.GetSheet(sheet_name))

    def list_fields(self) -> None:
        """List all fields' names in the document."""
        field_descriptions = self.com.GetFieldDescriptions() # IArrayOfFieldDescription
        for i in range(field_count := int(self.com.GetFieldCount())):
            print(f"{field_descriptions[i].Name!r}") # IFieldDescription.Name - print quotes around
        
    def get_field(self, field_name: str) -> object:
        """Get a field COM object by *field_name* and create a Field object from it."""
        return Field(self.com.GetField(field_name))
      
    def list_current_selections(self) -> None:
        """List all fields' selections currently affecting the document."""
        field_names = self.com.GetCurrentSelections().VarId
        selected_values = self.com.GetCurrentSelections().Selections
        if len(field_names) == 0:
            print("No filters are being applied.")
        else:
            for i in range(len(field_names)):    
                print(f">>> {field_names[i]!r}:")
                field_selections = selected_values[i].split(', ')  # selections are returned as a comma delimited str, so we have to split it
                for k in range(len(field_selections)):
                    if k == (len(field_selections) - 1):
                        print(f"{field_selections[k]!r}") # last item, don't use comma as a separator
                    else:
                        print(f"{field_selections[k]!r},")
                print()

    def get_straight_table_box_handles(self) -> tuple[object]:
        """
        Returns object handles for the current document's straight table,
        and to the list boxes containing its dimensions and expressions.

        Objects are returned as a tuple:
        (straightablebox, dimensions, expressions)
        """
        try:
            sh = self.get_sheet('Отчеты')
        except:
            sh = self.get_sheet('Универсальная таблица')
        if not sh: # if no sheet has been found
            raise ValueError('Не найден лист с универсальной таблицей.')

        # обращение к любому методу объекта StraightTableBox
        # 'Отчет по spaceman подробно 26.07.2021-09.08.2021'
        # (Документ 'BMK_Запасы_fusion', лист 'Универсальная таблица')
        # приводит к отказу приложения, поэтому вместо поиска нужного объекта
        # через цикл необходимо ссылаться на нужную таблицу
        # напрямую через ее индекс    
        if 'BMK_Запасы_fusion' in self.path:
            stboxes = sh.com.GetStraightTableBoxes()
            turnover_straight_table_box_index = 1 # изменить при необходимости
            stbox = stboxes[turnover_straight_table_box_index]
            stbox = StraightTableBox(stbox)
        elif not (stbox := sh.get_straight_table_box('Универсальная таблица')):
            stbox = sh.get_straight_table_box('Отчеты')

        dims = sh.get_list_box('Поля')
        exprs = sh.get_list_box('Выражения')

        return stbox, dims, exprs


class Field(object):
    
    """A field is used to filter the document."""
    def __init__(self, com: object) -> None:
        """Initialize the object."""
        # doc.com.Clear() # clear all fields to disable cross-filtering        
        self.com = com
        self.name = com.GetDescription().Name
        self.max_value_count = (
            com.GetValueCount(1) +  # 1 - selected - fields selected directly
            com.GetValueCount(2) +  # 2 - optional - all avilable values - all values when no values are selected, 0 when a value is selected, possible values when other fields are cross filtering the field
            com.GetValueCount(3) +  # 3 - deselected
            com.GetValueCount(4) +  # 4 - alternative - (=other than selected) 0 - when none are selected, count of possible values when a value is selected)
            com.GetValueCount(5)    # 5 - excluded -  all fields that are removed by other fields' cross-filtering
        )
       

    def __repr__(self) -> str:
        """Give some info about an object instance."""
        return f"<qlik.Field - '{self.name}', is_numeric={self.com.GetProperties().IsNumeric}, max_value_count={self.max_value_count}>"

    def list_values(self, vals: str ='optional', mode: str = 'text') -> None:
        """
        Create a list containing field values.
        Parameters:
            vals:
                'selected' - default 
                'optional'
                'deselected'
                'alternative'
                'excluded'
            
            mode:
                'text' - default
                'num' - numeric representation of field values
                        (e.g. ordinal number for dates)
                'bool' - returns true if field values are numeric
        """

        if vals == 'selected':
            arr = self.com.GetSelectedValues(self.com.GetValueCount(1))
        elif vals == 'optional':
            arr = self.com.GetOptionalValues(self.com.GetValueCount(2))
        elif vals == 'deselected':
            arr = self.com.GetDeselectedValues(self.com.GetValueCount(3))
        elif vals == 'alternative':
            arr = self.com.GetAlternativeValues(self.com.GetValueCount(4))
        elif vals == 'excluded':
            arr = self.com.GetExcludedValues(self.com.GetValueCount(5))
        
        ls = []

        for i in range(arr.Count):
            val = arr[i]
            if mode == 'text':
                val = val.Text
            elif mode == 'num':
                val = val.Number
            elif mode == 'bool':
                val = val.IsNumeric
            ls.append(val)
        print(f"Printing {vals.upper()!r} values:")
        print(ls)
        return ls

    
    def select_values(self, *args) -> None:
        """Select new values for the field."""

        values = list_args(args)
        self.com.Clear()
        
        is_numeric =  self.com.GetProperties().IsNumeric
        selection = self.com.GetNoValues()
        
        for i, val in enumerate(values):
            selection.Add() # add new value to our selection
            
            if is_numeric:
                selection(i).IsNumeric = True
                selection(i).Number = val
            else:
                selection(i).Text = val
        self.com.SelectValues(selection)

    def deselect_values(self, *args) -> None:
        """Deselect values for the field."""

        values_to_deselect = set(list_args(args))

        is_numeric =  self.com.GetProperties().IsNumeric
        
        selected_values = self.com.GetSelectedValues(self.com.GetValueCount(1))
        selected_values = {selected_values[i].Number if is_numeric else selected_values[i].Text for i in range(selected_values.Count)}

        values = selected_values - values_to_deselect
        selection = self.com.GetNoValues()
                
        for i, val in enumerate(values):
            selection.Add() # add new value to our selection
            
            if is_numeric:
                selection(i).IsNumeric = True
                selection(i).Number = val
            else:
                selection(i).Text = val
        self.com.SelectValues(selection)


class Sheet(object):
    """A sheet is one of the objects of a QlikView document."""

    def __init__(self, com: object) -> None:
        """Initialize the class."""
        self.com = com
        self.name = com.GetProperties().Name
    
    def __repr__(self) -> str:
        """Return the name of an object instance."""
        return f"<qlik.Sheet - '{self.name}'>"
    
    def list_straight_table_boxes(self) -> None:
        """Print the names of all straight table boxes in the sheet."""
        straight_table_boxes = self.com.GetStraightTableBoxes()
        for i in range(len(straight_table_boxes)):
            print(f"{straight_table_boxes[i].GetCaption().Name.v!r}")
    
    def get_straight_table_box(self, straight_table_box_name) -> None:
        """Get the straight table box object by name."""
        straight_table_boxes = self.com.GetStraightTableBoxes()
        for i in range(len(straight_table_boxes)):
            if straight_table_boxes[i].GetCaption().Name.v == straight_table_box_name:
                return StraightTableBox(straight_table_boxes[i])

    def list_list_boxes(self) -> None:
        """Print the names of all list boxes in the sheet."""
        list_boxes = self.com.GetListBoxes()
        for i in range(len(list_boxes)):
            print(f"{list_boxes[i].GetCaption().Name.v!r}")

    def get_list_box(self, list_box_name) -> None:
        """Get the list box object by name."""
        list_boxes = self.com.GetListBoxes()
        for i in range(len(list_boxes)):
            if list_boxes[i].GetCaption().Name.v == list_box_name:
                return ListBox(list_boxes[i])


class ListBox(object):
    """A list box is one of a Sheet's objects."""
    def __init__(self, com: object) -> None:
        """Initialize the class."""
        self.com = com
        self.name = com.GetCaption().Name.v

    def __repr__(self) -> str:
        """Return the name of the list box instance."""
        return f"<qlik.ListBox - '{self.name}'>"

    def get_field(self) -> object:
        """Return the field shown in the list box."""
        field = Field(self.com.GetField())
        return field


class StraightTableBox(object):
    """A straight table is one of a Sheet's objects."""
    def __init__(self, com: object) -> None:
        """Initialize the class."""
        self.com = com
        self.name = com.GetCaption().Name.v

    def __repr__(self) -> str:
        """Return the name of the straight table box instance."""
        return f"<qlik.StraightTableBox - '{self.name}'>"

    def export(self, path, sep: str = ';', append: bool = False) -> None:
        """Export the table to a file specified in path. *path* can be a string or a Path object."""
        if isinstance(path, str):
            path = Path(path)
        if path.suffix == '.csv':
            self.com.ExportEx(
                str(path), # convert the Path object to string
                1, # csv export mode
                append, # append mode
                sep # separator
            )
        elif path.suffix == '.xls':
             self.com.ExportEx(
                str(path),
                5, # Excel export mode
                append, # append mode
                sep # separator
            )
        else:
            raise TypeError("Can export only to .xls or .csv file format.")
    
    def to_df(
            self,
            from_ext: str = 'xls',
            skiprows=[1], # skip grand total row by default
            sep=';',
            **kwargs):
        """
        Export the straight table to a temp file
        and read it into a DataFrame object.
        Temp file is discarded after being read.

        `from_ext` specifies the format of a temporary file you are initially
        exporting to. 
        """
        tempfile_path: Path = gen_tempfile_name(ext=from_ext)
        self.export(path=tempfile_path, sep=sep, append=False)

        if from_ext == 'csv':
            df = pd.read_csv(
                tempfile_path,
                sep=sep,
                skiprows=skiprows, # skip totals row
                decimal=',',
                **kwargs
                # when types are mixed on import set dtypes explicitly
                # via the dtype kwarg
                # see low_memory section
                # @ https://pandas.pydata.org/docs/reference/api/pandas.read_csv.html
            ) 
        else:
            df = pd.read_excel(
                tempfile_path,
                skiprows=[1],
                **kwargs
            )
        
        tempfile_path.unlink()

        return df


