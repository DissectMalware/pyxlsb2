import warnings
from datetime import datetime, timedelta
from .workbook import Workbook
from .xlsbpackage import XlsbPackage

__version__ = '0.0.5'


def open_workbook(name, *args, **kwargs):
    """Opens the given workbook file path.

    Args:
        name (str): The name of the XLSB file to open.

    Returns:
        Workbook: The workbook instance for the given file name.

    Examples:
        This is typically the entrypoint to start working with an XLSB file:

        >>> from pyxlsb import open_workbook
        >>> with open_workbook('test_files/test.xlsb') as wb:
        ...     print(wb.sheets)
        ...
        ['Test']
    """
    return Workbook(XlsbPackage(name), *args, **kwargs)

