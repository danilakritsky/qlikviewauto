from .list_box import ListBox
from .straight_table_box import StraightTableBox

class Sheet:
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