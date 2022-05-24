from pathlib import Path

import pandas as pd
from bmk.functions import gen_tempfile_name


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

