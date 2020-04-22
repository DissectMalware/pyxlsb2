import sys
import os
import logging
import warnings
from datetime import datetime, timedelta
from . import recordtypes as rt
from .recordreader import RecordReader
from .stringtable import StringTable
from .styles import Styles
from .worksheet import Worksheet

if sys.version_info > (3,):
    basestring = (str, bytes)
    long = int

_MICROSECONDS_IN_DAY = 24 * 60 * 60 * 1000 * 1000


class Workbook(object):
    """The main Workbook class providing worksheets and other metadata.

    Attributes:
        sheets (list(str)): The worksheet names in this workbook.
    """

    def __init__(self, pkg):
        self._pkg = pkg
        self._parse()


    def __enter__(self):
        return self

    def __exit__(self, typ, value, traceback):
        self.close()

    def _parse(self):
        self.props = None
        self.sheets = list()
        self.stringtable = None
        self.styles = None

        workbook_rels = self._pkg.get_workbook_rels()
        with self._pkg.get_workbook_part() as f:
            for rectype, rec in RecordReader(f):
                if rectype == rt.WB_PROP:
                    self.props = rec
                elif rectype == rt.BUNDLE_SH:
                    rec.type, rec.loc = self.get_sheet_info(workbook_rels, rec.rId)
                    self.sheets.append(rec)
                elif rectype == rt.END_BUNDLE_SHS:
                    break

        ssfp = self._pkg.get_sharedstrings_part()
        if ssfp is not None:
            self.stringtable = StringTable(ssfp)

        stylesfp = self._pkg.get_styles_part()
        if stylesfp is not None:
            self.styles = Styles(stylesfp)

    def get_sheet_info(self, workbook_rels, rId):
        sheet_type = None
        sheet_path = None
        nsmap = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
        relationships = workbook_rels.findall('.//r:Relationship', namespaces=nsmap)

        for relationship in relationships:
            if relationship.attrib['Id'] == rId:
                sheet_path = relationship.attrib['Target']
                if relationship.attrib[
                    'Type'] == "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" or \
                        relationship.attrib[
                            'Type'] == 'http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet':
                    sheet_type = 'macrosheet'
                elif relationship.attrib[
                    'Type'] == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet":
                    sheet_type = 'worksheet'
                else:
                    sheet_type = 'unknown'
                break

        return sheet_type, sheet_path

    def get_workbook_rels(self):
        if not self._workbook_rels:
            workbook = self.get_xml_file('xl/_rels/workbook.xml.rels')
            self._workbook_rels = workbook
        return self._workbook_rels

    def get_sheet_by_index(self, idx, with_rels=False):
        """Get a worksheet by index.

        Args:
            idx (int): The index of the sheet.
            with_rels (bool, optional): If the relationships should be parsed, defaults to False

        Returns:
            Worksheet: The corresponding worksheet instance.

        Raises:
            IndexError: When the provided index is out of range.
        """
        if idx < 0 or idx >= len(self.sheets):
            raise IndexError('sheet index out of range')

        fp = self._pkg.get_worksheet_part(idx + 1)
        if with_rels:
            rels_fp = self._pkg.get_worksheet_rels(idx + 1)
        else:
            rels_fp = None
        return Worksheet(self, self.sheets[idx], fp, rels_fp)

    def get_sheet_by_name(self, name, with_rels=True):
        """Get a boundsheet by its name.

        Args:
            name (str): The name of the sheet.
            with_rels (bool, optional): If the relationships should be parsed, defaults to False

        Returns:
            Boundsheet: The corresponding worksheet instance.

        Raises:
            IndexError: When the provided name is invalid.
        """
        n = name.lower()
        loc = None
        target_sheet = None
        for sheet in self.sheets:
            if sheet.name.lower() == n:
                loc = sheet.loc
                target_sheet = sheet
                break

        sheet_fname = loc[loc.rfind('/')+1:]
        sheet_dir = loc[:loc.rfind('/')]
        fp = self._pkg.get_file('xl/'+loc)

        if with_rels:
            rels_fp = self._pkg.get_file('xl/{}/_rels/{}.rels'.format(sheet_dir, sheet_fname))
        else:
            rels_fp = None

        return Worksheet(self, target_sheet, fp, rels_fp)

    def get_shared_string(self, idx):
        """Get a string in the shared string table.

        Args:
            idx (int): The index of the string in the shared string table.

        Returns:
            str: The string at the index in table table or None if not found.
        """
        if self.stringtable is not None:
            return self.stringtable.get_string(idx)

    def convert_date(self, value):
        """Convert an Excel numeric date value to a ``datetime`` instance.

        Args:
            value (int or float): The Excel date value.

        Returns:
            datetime.datetime: The equivalent datetime instance or None when invalid.
        """
        if not isinstance(value, (int, long, float)):
            return None

        era = datetime(1904 if self.props.date1904 else 1900, 1, 1, tzinfo=None)
        timeoffset = timedelta(microseconds=long((value % 1) * _MICROSECONDS_IN_DAY))

        if int(value) == 0:
            return era + timeoffset

        if not self.props.date1904 and value >= 61:
            # According to Lotus 1-2-3, there is a Feb 29th 1900,
            # so we have to remove one extra day after that date
            dateoffset = timedelta(days=int(value) - 2)
        else:
            dateoffset = timedelta(days=int(value) - 1)

        return era + dateoffset + timeoffset

    def convert_time(self, value):
        """Convert an Excel date fraction time value to a ``time`` instance.

        Args:
            value (float): The Excel fractional day value.

        Returns:
            datetime.time: The equivalent time instance or None if invalid.
        """
        if not isinstance(value, (int, long, float)):
            return None
        return (datetime.min + timedelta(microseconds=long((value % 1) * _MICROSECONDS_IN_DAY))).time()

    def close(self):
        """Release the resources associated with this workbook."""
        self._pkg.close()
        if self.stringtable is not None:
            self.stringtable.close()
        if self.styles is not None:
            self.styles.close()
