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
                    formula_str = Formula.parse(cell.formula)
                    if formula_str.tokens:
                        try:
                            pass
                            print(formula_str.stringify(wb, cell.row_num, cell.col))
                        except NotImplementedError as exp:
                            print('ERROR({}) {}'.format(exp, str(cell)))
                        except Exception as e:
                            print('ERROR ' + str(cell)) + ' (%s)' % e
        d = time.time() - a
        print('Done! ({} seconds)'.format(d))
