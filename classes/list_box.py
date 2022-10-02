from .field import Field


class ListBox:
    """A list box is one of a Sheet's objects."""

    def __init__(self, com: object) -> None:
        """Initialize the class."""
        self.com = com
        self.name = com.GetCaption().Name.v

    def __repr__(self) -> str:
        """Return the name of the list box instance."""
        return f"<qlik.ListBox - '{self.name}'>"

    # TODO: make it a property
    def get_field(self) -> object:
        """Return the field shown in the list box."""
        field = Field(self.com.GetField())
        return field
