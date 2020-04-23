import sys
import time
from pyxlsb2 import open_workbook
from pyxlsb2.formula import Formula

a = time.time()
print('Opening workbook... ', end='', flush=True)
with open_workbook(sys.argv[1]) as wb:
    d = time.time() - a
    print('Done! ({} seconds)'.format(d))
    for s in wb.sheets:
        print('Reading sheet {}...\n'.format(s), end='', flush=True)
        a = time.time()
        with wb.get_sheet_by_name(s.name) as sheet:
            for row in sheet:
                for cell in row:
                    f = Formula.parse(cell.formula)
                    if f._tokens:
                        print(f.stringify(wb))
        d = time.time() - a
        print('Done! ({} seconds)'.format(d))
