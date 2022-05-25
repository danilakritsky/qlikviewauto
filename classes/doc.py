from .straight_table_box import StraightTableBox
from .field import Field
from .sheet import Sheet

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
            print('Trying to select sheet:')
            sh = self.get_sheet('Отчеты')
        except:
            sh = self.get_sheet('Универсальная таблица')
        if not sh: # if no sheet has been found
            raise ValueError('Не найден лист с универсальной таблицей.')

        # обращение к любому методу объекта StraightTableBox
        # 'Отчет по spaceman подробно 26.07.2021-09.08.2021'
        # (Документы 'BMK_Запасы_fusion', 'BWD_Запасы_fusion' и их архивные приложения, лист 'Универсальная таблица')
        # приводит к отказу приложения, поэтому вместо поиска нужного объекта
        # через цикл необходимо ссылаться на нужную таблицу
        # напрямую через ее индекс    
        if 'Запасы' in self.path:
            stboxes = sh.com.GetStraightTableBoxes()
            turnover_straight_table_box_index = 0 # изменить при необходимости
            print(f'Getting straight table box at index {turnover_straight_table_box_index}.')
            stbox = stboxes[turnover_straight_table_box_index]
            stbox = StraightTableBox(stbox)
        elif not (stbox := sh.get_straight_table_box('Универсальная таблица')):
            stbox = sh.get_straight_table_box('Отчеты')

        dims = sh.get_list_box('Поля')
        exprs = sh.get_list_box('Выражения')

        return stbox, dims, exprs