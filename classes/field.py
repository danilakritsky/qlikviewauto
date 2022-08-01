from bmk.functions import list_args


class Field:
    
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