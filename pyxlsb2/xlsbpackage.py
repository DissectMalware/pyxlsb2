import os
import shutil
from tempfile import TemporaryFile
from zipfile import ZipFile
from glob import fnmatch
from xml.etree import ElementTree

class ZipPackage(object):
    def __init__(self, name):
        self._zf_path = name
        self._zf = ZipFile(name, 'r')

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.close()

    def get_file(self, name):
        tf = TemporaryFile()
        try:
            with self._zf.open(name, 'r') as zf:
                shutil.copyfileobj(zf, tf)
            tf.seek(0, os.SEEK_SET)
            return tf
        except KeyError:
            tf.close()
            return None

    def get_files(self, file_name_filters=None):
        result = {}
        if not file_name_filters:
            file_name_filters = ['*']

        for i in self._zf.namelist():
            for filter in file_name_filters:
                if fnmatch.fnmatch(i, filter):
                    result[i] = self._zf.read(i)

        return result

    def get_xml_file(self, file_name):
        result = None
        files = self.get_files([file_name])
        if len(files) == 1:
            workbook_content = files[file_name].decode('utf_8')
            result = ElementTree.fromstring(workbook_content)
        return result

    def close(self):
        self._zf.close()


class XlsbPackage(ZipPackage):

    def get_workbook_part(self):
        return self.get_file('xl/workbook.bin')

    def get_workbook_rels(self):
        return self.get_xml_file('xl/_rels/workbook.bin.rels')

    def get_sharedstrings_part(self):
        return self.get_file('xl/sharedStrings.bin')

    def get_styles_part(self):
        return self.get_file('xl/styles.bin')

    def get_worksheet_part(self, idx):
        return self.get_file('xl/worksheets/sheet{}.bin'.format(idx))

    def get_worksheet_rels(self, idx):
        return self.get_file('xl/worksheets/_rels/sheet{}.bin.rels'.format(idx))
