import os
import shutil
from tempfile import TemporaryFile
from zipfile import ZipFile
from glob import fnmatch
from xml.etree import ElementTree
from io import BytesIO

class ZipPackage(object):
    def __init__(self, name):
        self._zf_path = name
        self._zf = ZipFile(name, 'r')

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.close()

    def get_file(self, name, in_mem=True):
        try:
            if in_mem:
                content = self._zf.read(name)
                tf = BytesIO(content)
            else:
                tf = TemporaryFile()
                with self._zf.open(name, 'r') as zf:
                    shutil.copyfileobj(zf, tf)
                tf.seek(0, os.SEEK_SET)
            return tf
        except KeyError:
            if not in_mem:
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
        ret = self.get_file('xl/workbook.bin')
        if not ret:
          ret = self.get_file('xl\\workbook.bin')
        return ret
    def get_workbook_rels(self):
        ret = self.get_xml_file('xl/_rels/workbook.bin.rels')
        if not ret:
            ret = self.get_xml_file('xl\\_rels\\workbook.bin.rels')
        return ret
    def get_sharedstrings_part(self):
        ret = self.get_file('xl/sharedStrings.bin')
        if not ret:
            ret = self.get_file('xl\\sharedStrings.bin')
        return ret
    def get_styles_part(self):
        ret = self.get_file('xl/styles.bin')
        if not ret:
            ret = self.get_file('xl\\styles.bin')
        return ret
    def get_worksheet_part(self, idx):
        ret = self.get_file('xl/worksheets/sheet{}.bin'.format(idx))
        if not ret:
            ret = self.get_file('xl\\worksheets\\sheet{}.bin'.format(idx))
        return ret
    def get_worksheet_rels(self, idx):
        ret = self.get_file('xl/worksheets/_rels/sheet{}.bin.rels'.format(idx))
        if not ret:
            ret = self.get_file('xl\\worksheets\\_rels\\sheet{}.bin.rels'.format(idx))
        return ret
