import logging

from comtypes import COMError

from .field import Field
from .sheet import Sheet


logger: logging.Logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
sh = logging.StreamHandler()
sh.setLevel(logging.DEBUG)
fmt = logging.Formatter(
    "{asctime} | {name} | {levelname} | {threadName} | {funcName} | {message}",
    style="{",
    datefmt="%Y-%m-%d %H:%M:%S",
)
sh.setFormatter(fmt)
logger.addHandler(sh)


class Doc:
    """A class representing a QlikView document."""

    def __init__(self, doc_path: str, com: object) -> object:
        """Initialize the class"""
        self.path: str = doc_path
        self.com: object = com  # @ https://stackoverflow.com/questions/625083/what-init-and-self-do-in-python
        self.com.Clear()

    def __repr__(self) -> str:
        """Give some info about an object instance."""
        return f"<qlik.Doc - '{self.path}'>"

    def list_sheets(self) -> list[str]:
        """List all sheets' names in the document."""
        return [
            f"{self.com.Sheets(str(i)).GetProperties().Name}"
            for i in range(self.com.NoOfSheets())
        ]

    def get_sheet(self, sheet_name: str) -> Sheet:
        """Get a sheet COM object by *name* and create a Sheet object from it."""
        return Sheet(self.com.GetSheet(sheet_name))

    def list_fields(self) -> list[str]:
        """List all fields' names in the document."""
        field_descriptions = self.com.GetFieldDescriptions()  # IArrayOfFieldDescription
        return [
            f"{field_descriptions[i].Name}"
            for i in range(int(self.com.GetFieldCount()))
        ]

    def get_field(self, field_name: str) -> Field:
        """Get a field COM object by *field_name* and create a Field object from it."""
        return Field(self.com.GetField(field_name))

    def list_current_selections(self) -> dict[str, list[str]]:
        """List all fields' selections currently affecting the document."""
        field_names: tuple[str] = self.com.GetCurrentSelections().VarId
        selected_values: tuple[str] = self.com.GetCurrentSelections().Selections
        selections: dict[str, list[str]] = {}
        if field_names:
            for i in range(len(field_names)):
                # selections are returned as a comma delimited str, so we have to split it
                selections[field_names[i]] = selected_values[i].split(", ")
        return selections
