import sys
from enum import Enum
from . import recordtypes as rt

if sys.version_info > (3,):
    xrange = range


class BasePtg(object):
    def __repr__(self):
        args = ('{}={}'.format(str(k), repr(v)) for k, v in self.__dict__.items())
        return '{}({})'.format(self.__class__.__name__, ', '.join(args))

    def is_classified(self):
        return False

    def stringify(self, tokens, workbook):
        return '#PTG{}!'.format(self.ptg)

    @classmethod
    def read(cls, reader, ptg):
        return cls()

    def write(self, writer):
        # TODO Eventually, some day
        pass





class ClassifiedPtg(BasePtg):
    def __init__(self, ptg, *args, **kwargs):
        super(ClassifiedPtg, self).__init__(*args, **kwargs)
        self.ptg = ptg

    @property
    def base_ptg(self):
        return ((self.ptg | 0x20) if self.ptg & 0x40 == 0x40 else self.ptg) & 0x3F

    def is_classified(self):
        return True

    def is_reference(self):
        return self.ptg & 0x20 == 0x20 and self.ptg & 0x40 == 0

    def is_value(self):
        return self.ptg & 0x20 == 0 and self.ptg & 0x40 == 0x40

    def is_array(self):
        return self.ptg & 0x60 == 0x60

    @classmethod
    def read(cls, reader, ptg):
        return cls(ptg)

    @staticmethod
    def cell_address(col, row, col_rel, row_rel):
        # A1 addressing
        col_name = ClassifiedPtg.convert_to_column_name(col+1)
        col_name = '$'+col_name if col_rel else col_name
        row_name = str(row+1)
        row_name = '$'+row_name if row_rel else row_name

        return col_name+row_name

    @staticmethod
    def convert_to_column_name(n):
        string = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string


class UnknownPtg(BasePtg):
    ptg = 0xFF

    def __init__(self, ptg, *args, **kwargs):
        super(UnknownPtg, self).__init__(*args, **kwargs)
        self.ptg = ptg

    def stringify(self, tokens, workbook):
        return '#UNK{}!'.format(self.ptg)

    @classmethod
    def read(cls, reader, ptg):
        return cls(ptg)


# Unary operators


class UPlusPtg(BasePtg):
    ptg = 0x12

    def stringify(self, tokens, workbook):
        return '+' + tokens.pop().stringify(tokens, workbook)


class UMinusPtg(BasePtg):
    ptg = 0x13

    def stringify(self, tokens, workbook):
        return '-' + tokens.pop().stringify(tokens, workbook)


class PercentPtg(BasePtg):
    ptg = 0x14

    def stringify(self, tokens, workbook):
        return tokens.pop().stringify(tokens, workbook) + '%'


# Binary operators


class AddPtg(BasePtg):
    ptg = 0x03

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '+' + b


class SubstractPtg(BasePtg):
    ptg = 0x04

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '-' + b


class MultiplyPtg(BasePtg):
    ptg = 0x05

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '*' + b


class DividePtg(BasePtg):
    ptg = 0x06

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '/' + b


class PowerPtg(BasePtg):
    ptg = 0x07

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '^' + b


class ConcatPtg(BasePtg):
    ptg = 0x08

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '&' + b


class LessPtg(BasePtg):
    ptg = 0x09

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '<' + b


class LessEqualPtg(BasePtg):
    ptg = 0x0A

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '<=' + b


class EqualPtg(BasePtg):
    ptg = 0x0B

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '=' + b


class GreaterEqualPtg(BasePtg):
    ptg = 0x0C

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '>=' + b


class GreaterPtg(BasePtg):
    ptg = 0x0D

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '>' + b


class NotEqualPtg(BasePtg):
    ptg = 0x0E

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + '<>' + b


class IntersectionPtg(BasePtg):
    ptg = 0x0F

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + ' ' + b


class UnionPtg(BasePtg):
    ptg = 0x10

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + ',' + b


class RangePtg(BasePtg):
    ptg = 0x11

    def stringify(self, tokens, workbook):
        b = tokens.pop().stringify(tokens, workbook)
        a = tokens.pop().stringify(tokens, workbook)
        return a + ':' + b


# Operands


class MissArgPtg(BasePtg):
    ptg = 0x16

    def stringify(self, tokens, workbook):
        return ''


class StringPtg(BasePtg):
    ptg = 0x17

    def __init__(self, value, *args, **kwargs):
        super(StringPtg, self).__init__(*args, **kwargs)
        self.value = value

    def stringify(self, tokens, workbook):
        return '"' + self.value.replace('"', '""') + '"'

    @classmethod
    def read(cls, reader, ptg):
        size = reader.read_short()
        value = reader.read_string(size=size)
        return cls(value)


class ErrorPtg(BasePtg):
    ptg = 0x1C

    def __init__(self, value, *args, **kwargs):
        super(ErrorPtg, self).__init__(*args, **kwargs)
        self.value = value

    def stringify(self, tokens, workbook):
        if self.value == 0x00:
            return '#NULL!'
        elif self.value == 0x07:
            return '#DIV/0!'
        elif self.value == 0x0F:
            return '#VALUE!'
        elif self.value == 0x17:
            return '#REF!'
        elif self.value == 0x1D:
            return '#NAME?'
        elif self.value == 0x24:
            return '#NUM!'
        elif self.value == 0x2A:
            return '#N/A'
        else:
            return '#ERR!'

    @classmethod
    def read(cls, reader, ptg):
        value = reader.read_byte()
        return cls(value)


class BooleanPtg(BasePtg):
    ptg = 0x1D

    def __init__(self, value, *args, **kwargs):
        super(BooleanPtg, self).__init__(*args, **kwargs)
        self.value = value

    def stringify(self, tokens, workbook):
        return 'TRUE' if self.value else 'FALSE'

    @classmethod
    def read(cls, reader, ptg):
        value = reader.read_bool()
        return cls(value)


class IntegerPtg(BasePtg):
    ptg = 0x1E

    def __init__(self, value, *args, **kwargs):
        super(IntegerPtg, self).__init__(*args, **kwargs)
        self.value = value

    def stringify(self, tokens, workbook):
        return str(self.value)

    @classmethod
    def read(cls, reader, ptg):
        value = reader.read_short()
        return cls(value)


class NumberPtg(BasePtg):
    ptg = 0x1F

    def __init__(self, value, *args, **kwargs):
        super(NumberPtg, self).__init__(*args, **kwargs)
        self.value = value

    def stringify(self, tokens, workbook):
        return str(self.value)

    @classmethod
    def read(cls, reader, ptg):
        value = reader.read_double()
        return cls(value)


class ArrayPtg(ClassifiedPtg):
    ptg = 0x20

    def __init__(self, cols, rows, values, *args, **kwargs):
        super(ArrayPtg, self).__init__(*args, **kwargs)
        self.cols = cols
        self.rows = rows
        self.values = values

    @classmethod
    def read(cls, reader, ptg):
        cols = reader.read_byte()
        if cols == 0:
            cols = 256
        rows = reader.read_short()
        values = list()
        for i in xrange(cols * rows):
            flag = reader.read_byte()
            value = None
            if flag == 0x01:
                value = reader.read_double()
            elif flag == 0x02:
                size = reader.read_short()
                value = reader.read_string(size=size)
            values.append(value)
        return cls(cols, rows, values, ptg)


class NamePtg(ClassifiedPtg):
    ptg = 0x23

    def __init__(self, idx, reserved, *args, **kwargs):
        super(NamePtg, self).__init__(*args, **kwargs)
        self.idx = idx
        self._reserved = reserved

    def stringify(self, tokens, workbook):
        defined = workbook.defined_names[workbook.list_names[self.idx - 1]]
        # return '%s (%s)' % (defined.name, defined.formula)
        return defined.formula if (defined.formula and '#NAME?' != defined.formula) else defined.name

    @classmethod
    def read(cls, reader, ptg):
        idx = reader.read_short()
        res = reader.read(2)  # Reserved
        return cls(idx, res, ptg)


class RefPtg(ClassifiedPtg):
    ptg = 0x24

    def __init__(self, row, col, row_rel, col_rel, *args, **kwargs):
        super(RefPtg, self).__init__(*args, **kwargs)
        self.row = row
        self.col = col
        self.row_rel = row_rel
        self.col_rel = col_rel


    def stringify(self, tokens, workbook):
        return self.cell_address(self.col, self.row, self.col_rel, self.row_rel)

    @classmethod
    def read(cls, reader, ptg):
        row = reader.read_int()
        col = reader.read_short()
        row_rel = col & 0x8000 == 0x8000
        col_rel = col & 0x4000 == 0x4000
        return cls(row, col & 0x3FFF, not row_rel, not col_rel, ptg)


class AreaPtg(ClassifiedPtg):
    ptg = 0x25

    def __init__(self, first_row, last_row, first_col, last_col, first_row_rel, last_row_rel, first_col_rel,
                 last_col_rel, *args, **kwargs):
        super(AreaPtg, self).__init__(*args, **kwargs)
        self.first_row = first_row
        self.last_row = last_row
        self.first_col = first_col
        self.last_col = last_col
        self.first_row_rel = first_row_rel
        self.last_row_rel = last_row_rel
        self.first_col_rel = first_col_rel
        self.last_col_rel = last_col_rel

    def stringify(self, tokens, workbook):
        first = self.cell_address(self.first_col, self.first_row, self.first_col_rel
                                  ,self.first_row_rel)
        last = self.cell_address(self.last_col, self.last_row, self.last_col_rel
                                  ,self.last_row_rel)
        return first + ':' + last

    @classmethod
    def read(cls, reader, ptg):
        r1 = reader.read_int()
        r2 = reader.read_int()
        c1 = reader.read_short()
        c2 = reader.read_short()
        r1rel = c1 & 0x8000 == 0x8000
        r2rel = c2 & 0x8000 == 0x8000
        c1rel = c1 & 0x4000 == 0x4000
        c2rel = c2 & 0x4000 == 0x4000
        return cls(r1, r2, c1 & 0x3FFF, c2 & 0x3FFF, not r1rel, not r2rel, not c1rel, not c2rel, ptg)


class MemAreaPtg(ClassifiedPtg):
    ptg = 0x26

    def __init__(self, reserved, rects, *args, **kwargs):
        super(ClassifiedPtg, self).__init__(*args, **kwargs)
        self._reserved = reserved
        self.rects = rects

    @classmethod
    def read(cls, reader, ptg):
        res = reader.read(4)  # Reserved
        subex_len = reader.read_short()
        rects = list()
        if subex_len:
            rect_count = reader.read_short()
            for i in xrange(rect_count):
                r1 = reader.read_int()
                r2 = reader.read_int()
                c1 = reader.read_short()
                c2 = reader.read_short()
                rects.append((r1, r2, c1, c2))
        return cls(res, rects, ptg)


class MemErrPtg(ClassifiedPtg):
    ptg = 0x27

    def __init__(self, reserved, subex, *args, **kwargs):
        super(MemErrPtg, self).__init__(*args, **kwargs)
        self._reserved = reserved
        self._subex = subex

    def stringify(self, tokens, workbook):
        return '#ERR!'

    @classmethod
    def read(cls, reader, ptg):
        res = reader.read(4)  # Reserved
        subex_len = reader.read_short()
        subex = reader.read(subex_len)
        return cls(res, subex, ptg)


class RefErrPtg(ClassifiedPtg):
    ptg = 0x2A

    def __init__(self, reserved, *args, **kwargs):
        super(RefErrPtg, self).__init__(*args, **kwargs)
        self._reserved = reserved

    def stringify(self, tokens, workbook):
        return '#REF!'

    @classmethod
    def read(cls, reader, ptg):
        res = reader.read(6)  # Reserved
        return cls(res, ptg)


class AreaErrPtg(ClassifiedPtg):
    ptg = 0x2B

    def __init__(self, reserved, *args, **kwargs):
        super(AreaErrPtg, self).__init__(*args, **kwargs)
        self._reserved = reserved

    def stringify(self, tokens, workbook):
        return '#REF!'

    @classmethod
    def read(cls, reader, ptg):
        res = reader.read(12)  # Reserved
        return cls(res, ptg)


class RefNPtg(ClassifiedPtg):
    ptg = 0x2C

    def __init__(self, row, col, row_rel, col_rel, *args, **kwargs):
        super(RefNPtg, self).__init__(*args, **kwargs)
        self.row = row
        self.col = col
        self.row_rel = row_rel
        self.col_rel = col_rel

    def stringify(self, tokens, workbook):
        return self.cell_address(self.col, self.row, self.col_rel, self.row_rel)

    @classmethod
    def read(cls, reader, ptg):
        row = reader.read_int()
        col = reader.read_short()
        row_rel = col & 0x8000 == 0x8000
        col_rel = col & 0x4000 == 0x4000
        return cls(row, col & 0x3FFF, not row_rel, not col_rel, ptg)


class AreaNPtg(ClassifiedPtg):
    ptg = 0x2D

    def __init__(self, first_row, last_row, first_col, last_col, first_row_rel, last_row_rel, first_col_rel,
                 last_col_rel, *args, **kwargs):
        super(AreaNPtg, self).__init__(*args, **kwargs)
        self.first_row = first_row
        self.last_row = last_row
        self.first_col = first_col
        self.last_col = last_col
        self.first_row_rel = first_row_rel
        self.last_row_rel = last_row_rel
        self.first_col_rel = first_col_rel
        self.last_col_rel = last_col_rel

    def stringify(self, tokens, workbook):
        first = self.cell_address(self.first_col, self.first_row, self.first_col_rel
                                  , self.first_row_rel)
        last = self.cell_address(self.last_col, self.last_row, self.last_col_rel
                                 , self.last_row_rel)
        return first + ':' + last

    @classmethod
    def read(cls, reader, ptg):
        r1 = reader.read_int()
        r2 = reader.read_int()
        c1 = reader.read_short()
        c2 = reader.read_short()
        r1_rel = c1 & 0x8000 == 0x8000
        r2_rel = c2 & 0x8000 == 0x8000
        c1_rel = c1 & 0x4000 == 0x4000
        c2_rel = c2 & 0x4000 == 0x4000
        return cls(r1, r2, c1 & 0x3FFF, c2 & 0x3FFF, not r1_rel, not r2_rel, not c1_rel, not c2_rel, ptg)


class NameXPtg(ClassifiedPtg):
    ptg = 0x39

    def __init__(self, sheet_idx, name_idx, reserved, *args, **kwargs):
        super(NameXPtg, self).__init__(*args, **kwargs)
        self.sheet_idx = sheet_idx
        self.name_idx = name_idx
        self._reserved = reserved

    # FIXME: We need to read names to stringify this

    @classmethod
    def read(cls, reader, ptg):
        sheet_idx = reader.read_short()
        name_idx = reader.read_short()
        res = reader.read(2)  # Reserved
        return cls(sheet_idx, name_idx, res, ptg)


class Ref3dPtg(ClassifiedPtg):
    ptg = 0x3A

    def __init__(self, extern_sheet_id, row, col, row_rel, col_rel, *args, **kwargs):
        super(Ref3dPtg, self).__init__(*args, **kwargs)
        self.extern_sheet_idx = extern_sheet_id
        self.row = row
        self.col = col
        self.row_rel = row_rel
        self.col_rel = col_rel

    def stringify(self, tokens, workbook):
        if (not self.col_rel) and (self.col & 0x2000 == 0x2000):
            self.col -= 16384
        if (not self.row_rel) and (self.row & 0x80000 == 0x80000):
            self.row -= 1048576
        cell_add = self.cell_address(self.col, self.row, self.col_rel, self.row_rel)

        supporting_link, first_sheet_idx, last_sheet_idx = workbook.resolve_extern_sheet_id(self.extern_sheet_idx)
        address = None
        if supporting_link.brt in rt._SUP_LINK_TYPES:
            if first_sheet_idx == last_sheet_idx and first_sheet_idx >= 0:
                address = "'{}'!{}".format(workbook.sheets[first_sheet_idx].name, cell_add)
            elif first_sheet_idx == last_sheet_idx and first_sheet_idx == -1:
                address = "'{}'!{}".format(workbook.sheets[last_sheet_idx].name, cell_add)
            elif first_sheet_idx == last_sheet_idx and first_sheet_idx == -2:
                address = cell_add

        if address is None:
            print("Ref3dPtg External Address Not Supported {0} {1} {2}".format(cell_add, first_sheet_idx, last_sheet_idx))
        #    raise NotImplementedError('External address not supported')

        return address

    @classmethod
    def read(cls, reader, ptg):
        sheet_extern_idx = reader.read_short()
        row = reader.read_int()
        col = reader.read_short()
        row_rel = col & 0x8000 == 0x8000
        col_rel = col & 0x4000 == 0x4000
        return cls(sheet_extern_idx, row, col & 0x3FFF, not row_rel, not col_rel, ptg)

    def cell_address(self, col, row, col_rel, row_rel):
        # External reference, address is already computed
        # if not (col_rel and row_rel):
        #     col_name = 'C' + str(col + 1) if col_rel else 'C[%s]' % str(col)
        #     row_name = 'R' + str(row + 1) if row_rel else 'R[%s]' % str(row)
        #     res = row_name + col_name
        # else:
        #     col_name = '$' + self.convert_to_column_name(col + 1)
        #     row_name = '$' + str(row + 1)
        #     res = col_name + row_name

        col_name = '$' + self.convert_to_column_name(col + 1) if col_rel else self.convert_to_column_name(col + 1)
        row_name = '$' + str(row + 1) if row_rel else str(row + 1)
        res = col_name + row_name


        return res


class Area3dPtg(ClassifiedPtg):
    ptg = 0x3B

    def __init__(self, extern_sheet_idx, first_row, last_row, first_col, last_col, first_row_rel, last_row_rel, first_col_rel,
                 last_col_rel, *args, **kwargs):
        super(Area3dPtg, self).__init__(*args, **kwargs)
        self.extern_sheet_idx = extern_sheet_idx
        self.first_row = first_row
        self.last_row = last_row
        self.first_col = first_col
        self.last_col = last_col
        self.first_row_rel = first_row_rel
        self.last_row_rel = last_row_rel
        self.first_col_rel = first_col_rel
        self.last_col_rel = last_col_rel

    def stringify(self, tokens, workbook):
        first = self.cell_address(self.first_col, self.first_row, self.first_col_rel
                                  ,self.first_row_rel)
        last = self.cell_address(self.last_col, self.last_row, self.last_col_rel
                                  ,self.last_row_rel)
        supporting_link, first_sheet_idx, last_sheet_idx = workbook.resolve_extern_sheet_id(self.extern_sheet_idx)

        cell_add_first = self.cell_address(self.first_col, self.first_row, self.first_col_rel, self.first_row_rel)
        cell_add_last = self.cell_address(self.last_col, self.last_row, self.last_col_rel, self.last_row_rel)

        address = None
        if supporting_link.brt in rt._SUP_LINK_TYPES:
            if first_sheet_idx == last_sheet_idx and first_sheet_idx >= 0:
                address = "'{}'!{}".format(workbook.sheets[first_sheet_idx].name , first + ':' + last)

        if address is None:
            print("AreaPtg External Address Not Supported {0}:{1} {2} {3}".format(cell_add_first, cell_add_last, first_sheet_idx, last_sheet_idx))
            #raise NotImplementedError('External address not supported')

        return address

    @classmethod
    def read(cls, reader, ptg):
        XtiIndex = reader.read_short()
        r1 = reader.read_int()
        r2 = reader.read_int()
        c1 = reader.read_short()
        c2 = reader.read_short()
        r1_rel = c1 & 0x8000 == 0x8000
        r2_rel = c2 & 0x8000 == 0x8000
        c1_rel = c1 & 0x4000 == 0x4000
        c2_rel = c2 & 0x4000 == 0x4000
        return cls(XtiIndex, r1, r2, c1 & 0x3FFF, c2 & 0x3FFF, not r1_rel, not r2_rel, not c1_rel, not c2_rel, ptg)


class RefErr3dPtg(ClassifiedPtg):
    ptg = 0x3C

    def __init__(self, reserved, *args, **kwargs):
        super(RefErr3dPtg, self).__init__(*args, **kwargs)
        self._reserved = reserved

    def stringify(self, tokens, workbook):
        return '#REF!'

    @classmethod
    def read(cls, reader, ptg):
        res = reader.read(8)  # Reserved
        return cls(res, ptg)


class AreaErr3dPtg(ClassifiedPtg):
    ptg = 0x3D

    def __init__(self, reserved, *args, **kwargs):
        super(AreaErr3dPtg, self).__init__(*args, **kwargs)

    def stringify(self, tokens, workbook):
        return '#REF!'

    @classmethod
    def read(cls, reader, ptg):
        res = reader.read(14)  # Reserved
        return cls(res, ptg)


# Control


class ExpPtg(BasePtg):
    ptg = 0x01

    def __init__(self, row, col, *args, **kwargs):
        super(ExpPtg, self).__init__(*args, **kwargs)
        self.row = row
        self.col = col

    # FIXME: We need a workbook that supports direct cell referencing to stringify this

    @classmethod
    def read(cls, reader, ptg):
        row = reader.read_int()
        col = reader.read_short()
        return cls(row, col)


class TablePtg(BasePtg):
    ptg = 0x02

    def __init__(self, row, col, *args, **kwargs):
        super(TablePtg, self).__init__(*args, **kwargs)
        self.row = row
        self.col = col

    # FIXME: We need a workbook that supports tables to stringify this

    @classmethod
    def read(cls, reader, ptg):
        row = reader.read_int()
        col = reader.read_short()
        return cls(row, col)


class ParenPtg(BasePtg):
    ptg = 0x15

    def stringify(self, tokens, workbook):
        return '(' + tokens.pop().stringify(tokens, workbook) + ')'


class AttrPtg(BasePtg):
    ptg = 0x19

    def __init__(self, flags, data, *args, **kwargs):
        super(AttrPtg, self).__init__(*args, **kwargs)
        self.flags = flags
        self.data = data

    @property
    def attr_semi(self):
        return self.flags & 0x01 == 0x01

    @property
    def attr_if(self):
        return self.flags & 0x02 == 0x02

    @property
    def attr_choose(self):
        return self.flags & 0x04 == 0x04

    @property
    def attr_goto(self):
        return self.flags & 0x08 == 0x08

    @property
    def attr_sum(self):
        return self.flags & 0x10 == 0x10

    @property
    def attr_baxcel(self):
        return self.flags & 0x20 == 0x20

    @property
    def attr_space(self):
        return self.flags & 0x40 == 0x40

    def stringify(self, tokens, workbook):
        spaces = ''
        if self.attr_space:
            if self.data & 0x00FF == 0x00 or self.data & 0x00FF == 0x06:
                spaces = ' ' * (self.data >> 1)
            elif self.data & 0x00FF == 0x01:
                spaces = '\n' * (self.data >> 1)
        if tokens:
            spaces +=  tokens.pop().stringify(tokens, workbook)
        return spaces

    @classmethod
    def read(cls, reader, ptg):
        flags = reader.read_byte()
        data = reader.read_short()
        return cls(flags, data)


class MemNoMemPtg(ClassifiedPtg):
    ptg = 0x28

    def __init__(self, reserved, subex, *args, **kwargs):
        super(MemNoMemPtg, self).__init__(*args, **kwargs)
        self._reserved = reserved
        self._subex = subex

    @classmethod
    def read(cls, reader, ptg):
        res = reader.read(4)  # Reserved
        subex_len = reader.read_short()
        subex = reader.read(subex_len)
        return cls(res, subex, ptg)


class MemFuncPtg(ClassifiedPtg):
    ptg = 0x29

    def __init__(self, subex, *args, **kwargs):
        super(MemFuncPtg, self).__init__(*args, **kwargs)
        self._subex = subex

    @classmethod
    def read(cls, reader, ptg):
        subex_len = reader.read_short()
        subex = reader.read(subex_len)
        return cls(subex, ptg)


class MemAreaNPtg(ClassifiedPtg):
    ptg = 0x2E

    def __init__(self, subex, *args, **kwargs):
        super(MemAreaNPtg, self).__init__(*args, **kwargs)
        self._subex = subex

    @classmethod
    def read(cls, reader, ptg):
        subex_len = reader.read_short()
        subex = reader.skip(subex_len)
        return cls(subex, ptg)


class MemNoMemNPtg(ClassifiedPtg):
    ptg = 0x2F

    def __init__(self, subex, *args, **kwargs):
        super(MemNoMemNPtg, self).__init__(*args, **kwargs)
        self._subex = subex

    @classmethod
    def read(cls, reader, ptg):
        subex_len = reader.read_short()
        subex = reader.read(subex_len)
        return cls(subex, ptg)


# Func operators


class FuncPtg(ClassifiedPtg):
    ptg = 0x21

    def __init__(self, idx, *args, **kwargs):
        super(FuncPtg, self).__init__(*args, **kwargs)
        self.idx = idx

    # FIXME: We need a workbook to stringify this (most likely)

    @classmethod
    def read(cls, reader, ptg):
        idx = reader.read_short()
        return cls(idx, ptg)

    def stringify(self, tokens, workbook):
        args = list()

        if len(function_names[self.idx])>1:
            args_no = function_names[self.idx][1]
            for i in xrange(args_no):
                arg = tokens.pop().stringify(tokens, workbook)
                args.append(arg)

        return '{}({})'.format(function_names[self.idx][0], ', '.join(reversed(args)))


class FuncVarPtg(ClassifiedPtg):
    ptg = 0x22

    def __init__(self, idx, argc, prompt, ce, *args, **kwargs):
        super(FuncVarPtg, self).__init__(*args, **kwargs)
        self.idx = idx
        self.argc = argc
        self.prompt = prompt
        self.ce = ce

    def stringify(self, tokens, workbook):
        if self.idx == 255:  # UserDefinedFunction
            function_name = tokens[0].stringify(tokens, workbook).strip()
            del tokens[0]
            self.argc -= 1
        else:
            function_name = function_names[self.idx][0]

        args = list()
        for i in xrange(self.argc):
            arg = tokens.pop().stringify(tokens, workbook).strip()
            args.append(arg)

        return '{}({})'.format(function_name, ', '.join(reversed(args)))

    @classmethod
    def read(cls, reader, ptg):
        argc = reader.read_byte()
        idx = reader.read_short()
        return cls(idx, argc & 0x7F, argc & 0x80 == 0x80, idx & 0x8000 == 0x8000, ptg)

function_names = {
    # https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/00b5dd7d-51ca-4938-b7b7-483fe0e5933b
    # Newer (more functions): https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/90a52fcb-ce63-497f-a3d3-173c42d82242
    0x0000: ('COUNT',),
    0x0001: ('IF',),
    0x0002: ('ISNA', 1),
    0x0003: ('ISERROR', 1),
    0x0004: ('SUM',),
    0x0005: ('AVERAGE',),
    0x0006: ('MIN',),
    0x0007: ('MAX',),
    0x0008: ('ROW',),
    0x0009: ('COLUMN',),
    0x000A: ('NA',),
    0x000B: ('NPV',),
    0x000C: ('STDEV',),
    0x000D: ('DOLLAR',),
    0x000E: ('FIXED',),
    0x000F: ('SIN', 1),
    0x0010: ('COS', 1),
    0x0011: ('TAN', 1),
    0x0012: ('ATAN', 1),
    0x0013: ('PI',),
    0x0014: ('SQRT', 1),
    0x0015: ('EXP', 1),
    0x0016: ('LN', 1),
    0x0017: ('LOG10', 1),
    0x0018: ('ABS', 1),
    0x0019: ('INT', 1),
    0x001A: ('SIGN', 1),
    0x001B: ('ROUND', 2),
    0x001C: ('LOOKUP',),
    0x001D: ('INDEX',),
    0x001E: ('REPT', 2),
    0x001F: ('MID', 3),
    0x0020: ('LEN', 1),
    0x0021: ('VALUE', 1),
    0x0022: ('TRUE',),
    0x0023: ('FALSE',),
    0x0024: ('AND',),
    0x0025: ('OR',),
    0x0026: ('NOT', 1),
    0x0027: ('MOD', 2),
    0x0028: ('DCOUNT', 3),
    0x0029: ('DSUM', 3),
    0x002A: ('DAVERAGE', 3),
    0x002B: ('DMIN', 3),
    0x002C: ('DMAX', 3),
    0x002D: ('DSTDEV', 3),
    0x002E: ('VAR',),
    0x002F: ('DVAR', 3),
    0x0030: ('TEXT', 2),
    0x0031: ('LINEST',),
    0x0032: ('TREND',),
    0x0033: ('LOGEST',),
    0x0034: ('GROWTH',),
    0x0035: ('GOTO', 1),
    0x0036: ('HALT',),
    0x0037: ('RETURN',),
    0x0038: ('PV',),
    0x0039: ('FV',),
    0x003A: ('NPER',),
    0x003B: ('PMT',),
    0x003C: ('RATE',),
    0x003D: ('MIRR', 3),
    0x003E: ('IRR',),
    0x003F: ('RAND',),
    0x0040: ('MATCH',),
    0x0041: ('DATE', 3),
    0x0042: ('TIME', 3),
    0x0043: ('DAY', 1),
    0x0044: ('MONTH', 1),
    0x0045: ('YEAR', 1),
    0x0046: ('WEEKDAY',),
    0x0047: ('HOUR', 1),
    0x0048: ('MINUTE', 1),
    0x0049: ('SECOND', 1),
    0x004A: ('NOW',),
    0x004B: ('AREAS', 1),
    0x004C: ('ROWS', 1),
    0x004D: ('COLUMNS', 1),
    0x004E: ('OFFSET',),
    0x004F: ('ABSREF', 2),
    0x0050: ('RELREF', 2),
    0x0051: ('ARGUMENT',),
    0x0052: ('SEARCH',),
    0x0053: ('TRANSPOSE', 1),
    0x0054: ('ERROR',),
    0x0055: ('STEP',),
    0x0056: ('TYPE', 1),
    0x0057: ('ECHO',),
    0x0058: ('SET.NAME',),
    0x0059: ('CALLER',),
    0x005A: ('DEREF', 1),
    0x005B: ('WINDOWS',),
    0x005C: ('SERIES',),
    0x005D: ('DOCUMENTS',),
    0x005E: ('ACTIVE.CELL',),
    0x005F: ('SELECTION',),
    0x0060: ('RESULT',),
    0x0061: ('ATAN2', 2),
    0x0062: ('ASIN', 1),
    0x0063: ('ACOS', 1),
    0x0064: ('CHOOSE',),
    0x0065: ('HLOOKUP',),
    0x0066: ('VLOOKUP',),
    0x0067: ('LINKS',),
    0x0068: ('INPUT',),
    0x0069: ('ISREF', 1),
    0x006A: ('GET.FORMULA', 1),
    0x006B: ('GET.NAME',),
    0x006C: ('SET.VALUE', 2),
    0x006D: ('LOG',),
    0x006E: ('EXEC',),
    0x006F: ('CHAR', 1),
    0x0070: ('LOWER', 1),
    0x0071: ('UPPER', 1),
    0x0072: ('PROPER', 1),
    0x0073: ('LEFT',),
    0x0074: ('RIGHT',),
    0x0075: ('EXACT', 2),
    0x0076: ('TRIM', 1),
    0x0077: ('REPLACE', 4),
    0x0078: ('SUBSTITUTE',),
    0x0079: ('CODE', 1),
    0x007A: ('NAMES',),
    0x007B: ('DIRECTORY',),
    0x007C: ('FIND',),
    0x007D: ('CELL',),
    0x007E: ('ISERR', 1),
    0x007F: ('ISTEXT', 1),
    0x0080: ('ISNUMBER', 1),
    0x0081: ('ISBLANK', 1),
    0x0082: ('T', 1),
    0x0083: ('N', 1),
    0x0084: ('FOPEN',),
    0x0085: ('FCLOSE', 1),
    0x0086: ('FSIZE', 1),
    0x0087: ('FREADLN', 1),
    0x0088: ('FREAD', 2),
    0x0089: ('FWRITELN', 2),
    0x008A: ('FWRITE', 2),
    0x008B: ('FPOS',),
    0x008C: ('DATEVALUE', 1),
    0x008D: ('TIMEVALUE', 1),
    0x008E: ('SLN', 3),
    0x008F: ('SYD', 4),
    0x0090: ('DDB',),
    0x0091: ('GET.DEF',),
    0x0092: ('REFTEXT',),
    0x0093: ('TEXTREF',),
    0x0094: ('INDIRECT',),
    0x0095: ('REGISTER',),
    0x0096: ('CALL',),
    0x0097: ('ADD.BAR',),
    0x0098: ('ADD.MENU',),
    0x0099: ('ADD.COMMAND',),
    0x009A: ('ENABLE.COMMAND',),
    0x009B: ('CHECK.COMMAND',),
    0x009C: ('RENAME.COMMAND',),
    0x009D: ('SHOW.BAR',),
    0x009E: ('DELETE.MENU',),
    0x009F: ('DELETE.COMMAND',),
    0x00A0: ('GET.CHART.ITEM',),
    0x00A1: ('DIALOG.BOX', 1),
    0x00A2: ('CLEAN', 1),
    0x00A3: ('MDETERM', 1),
    0x00A4: ('MINVERSE', 1),
    0x00A5: ('MMULT', 2),
    0x00A6: ('FILES',),
    0x00A7: ('IPMT',),
    0x00A8: ('PPMT',),
    0x00A9: ('COUNTA',),
    0x00AA: ('CANCEL.KEY',),
    0x00AB: ('FOR',),
    0x00AC: ('WHILE', 1),
    0x00AD: ('BREAK',),
    0x00AE: ('NEXT',),
    0x00AF: ('INITIATE', 2),
    0x00B0: ('REQUEST', 2),
    0x00B1: ('POKE', 3),
    0x00B2: ('EXECUTE', 2),
    0x00B3: ('TERMINATE', 1),
    0x00B4: ('RESTART',),
    0x00B5: ('HELP',),
    0x00B6: ('GET.BAR',),
    0x00B7: ('PRODUCT',),
    0x00B8: ('FACT', 1),
    0x00B9: ('GET.CELL',),
    0x00BA: ('GET.WORKSPACE', 1),
    0x00BB: ('GET.WINDOW',),
    0x00BC: ('GET.DOCUMENT',),
    0x00BD: ('DPRODUCT', 3),
    0x00BE: ('ISNONTEXT', 1),
    0x00BF: ('GET.NOTE',),
    0x00C0: ('NOTE',),
    0x00C1: ('STDEVP',),
    0x00C2: ('VARP',),
    0x00C3: ('DSTDEVP', 3),
    0x00C4: ('DVARP', 3),
    0x00C5: ('TRUNC',),
    0x00C6: ('ISLOGICAL', 1),
    0x00C7: ('DCOUNTA', 3),
    0x00C8: ('DELETE.BAR', 1),
    0x00C9: ('UNREGISTER', 1),
    0x00CC: ('USDOLLAR',),
    0x00CD: ('FINDB',),
    0x00CE: ('SEARCHB',),
    0x00CF: ('REPLACEB', 4),
    0x00D0: ('LEFTB',),
    0x00D1: ('RIGHTB',),
    0x00D2: ('MIDB', 3),
    0x00D3: ('LENB', 1),
    0x00D4: ('ROUNDUP', 2),
    0x00D5: ('ROUNDDOWN', 2),
    0x00D6: ('ASC', 1),
    0x00D7: ('DBCS', 1),
    0x00D8: ('RANK',),
    0x00DB: ('ADDRESS',),
    0x00DC: ('DAYS360',),
    0x00DD: ('TODAY',),
    0x00DE: ('VDB',),
    0x00DF: ('ELSE',),
    0x00E0: ('ELSE.IF', 1),
    0x00E1: ('END.IF',),
    0x00E2: ('FOR.CELL',),
    0x00E3: ('MEDIAN',),
    0x00E4: ('SUMPRODUCT',),
    0x00E5: ('SINH', 1),
    0x00E6: ('COSH', 1),
    0x00E7: ('TANH', 1),
    0x00E8: ('ASINH', 1),
    0x00E9: ('ACOSH', 1),
    0x00EA: ('ATANH', 1),
    0x00EB: ('DGET', 3),
    0x00EC: ('CREATE.OBJECT',),
    0x00ED: ('VOLATILE',),
    0x00EE: ('LAST.ERROR',),
    0x00EF: ('CUSTOM.UNDO',),
    0x00F0: ('CUSTOM.REPEAT',),
    0x00F1: ('FORMULA.CONVERT',),
    0x00F2: ('GET.LINK.INFO',),
    0x00F3: ('TEXT.BOX',),
    0x00F4: ('INFO', 1),
    0x00F5: ('GROUP',),
    0x00F6: ('GET.OBJECT',),
    0x00F7: ('DB',),
    0x00F8: ('PAUSE',),
    0x00FB: ('RESUME',),
    0x00FC: ('FREQUENCY', 2),
    0x00FD: ('ADD.TOOLBAR',),
    0x00FE: ('DELETE.TOOLBAR', 1),
    0x00FF: ('UserDefinedFunction',),
    0x0100: ('RESET.TOOLBAR', 1),
    0x0101: ('EVALUATE', 1),
    0x0102: ('GET.TOOLBAR',),
    0x0103: ('GET.TOOL',),
    0x0104: ('SPELLING.CHECK',),
    0x0105: ('ERROR.TYPE', 1),
    0x0106: ('APP.TITLE',),
    0x0107: ('WINDOW.TITLE',),
    0x0108: ('SAVE.TOOLBAR',),
    0x0109: ('ENABLE.TOOL', 3),
    0x010A: ('PRESS.TOOL', 3),
    0x010B: ('REGISTER.ID',),
    0x010C: ('GET.WORKBOOK',),
    0x010D: ('AVEDEV',),
    0x010E: ('BETADIST',),
    0x010F: ('GAMMALN', 1),
    0x0110: ('BETAINV',),
    0x0111: ('BINOMDIST', 4),
    0x0112: ('CHIDIST', 2),
    0x0113: ('CHIINV', 2),
    0x0114: ('COMBIN', 2),
    0x0115: ('CONFIDENCE', 3),
    0x0116: ('CRITBINOM', 3),
    0x0117: ('EVEN', 1),
    0x0118: ('EXPONDIST', 3),
    0x0119: ('FDIST', 3),
    0x011A: ('FINV', 3),
    0x011B: ('FISHER', 1),
    0x011C: ('FISHERINV', 1),
    0x011D: ('FLOOR', 2),
    0x011E: ('GAMMADIST', 4),
    0x011F: ('GAMMAINV', 3),
    0x0120: ('CEILING', 2),
    0x0121: ('HYPGEOMDIST', 4),
    0x0122: ('LOGNORMDIST', 3),
    0x0123: ('LOGINV', 3),
    0x0124: ('NEGBINOMDIST', 3),
    0x0125: ('NORMDIST', 4),
    0x0126: ('NORMSDIST', 1),
    0x0127: ('NORMINV', 3),
    0x0128: ('NORMSINV', 1),
    0x0129: ('STANDARDIZE', 3),
    0x012A: ('ODD', 1),
    0x012B: ('PERMUT', 2),
    0x012C: ('POISSON', 3),
    0x012D: ('TDIST', 3),
    0x012E: ('WEIBULL', 4),
    0x012F: ('SUMXMY2', 2),
    0x0130: ('SUMX2MY2', 2),
    0x0131: ('SUMX2PY2', 2),
    0x0132: ('CHITEST', 2),
    0x0133: ('CORREL', 2),
    0x0134: ('COVAR', 2),
    0x0135: ('FORECAST', 3),
    0x0136: ('FTEST', 2),
    0x0137: ('INTERCEPT', 2),
    0x0138: ('PEARSON', 2),
    0x0139: ('RSQ', 2),
    0x013A: ('STEYX', 2),
    0x013B: ('SLOPE', 2),
    0x013C: ('TTEST', 4),
    0x013D: ('PROB',),
    0x013E: ('DEVSQ',),
    0x013F: ('GEOMEAN',),
    0x0140: ('HARMEAN',),
    0x0141: ('SUMSQ',),
    0x0142: ('KURT',),
    0x0143: ('SKEW',),
    0x0144: ('ZTEST',),
    0x0145: ('LARGE', 2),
    0x0146: ('SMALL', 2),
    0x0147: ('QUARTILE', 2),
    0x0148: ('PERCENTILE', 2),
    0x0149: ('PERCENTRANK',),
    0x014A: ('MODE',),
    0x014B: ('TRIMMEAN', 2),
    0x014C: ('TINV', 2),
    0x014E: ('MOVIE.COMMAND',),
    0x014F: ('GET.MOVIE',),
    0x0150: ('CONCATENATE',),
    0x0151: ('POWER', 2),
    0x0152: ('PIVOT.ADD.DATA',),
    0x0153: ('GET.PIVOT.TABLE',),
    0x0154: ('GET.PIVOT.FIELD',),
    0x0155: ('GET.PIVOT.ITEM',),
    0x0156: ('RADIANS', 1),
    0x0157: ('DEGREES', 1),
    0x0158: ('SUBTOTAL',),
    0x0159: ('SUMIF',),
    0x015A: ('COUNTIF', 2),
    0x015B: ('COUNTBLANK', 1),
    0x015C: ('SCENARIO.GET',),
    0x015D: ('OPTIONS.LISTS.GET', 1),
    0x015E: ('ISPMT', 4),
    0x015F: ('DATEDIF', 3),
    0x0160: ('DATESTRING', 1),
    0x0161: ('NUMBERSTRING', 2),
    0x0162: ('ROMAN',),
    0x0163: ('OPEN.DIALOG',),
    0x0164: ('SAVE.DIALOG',),
    0x0165: ('VIEW.GET',),
    0x0166: ('GETPIVOTDATA',),
    0x0167: ('HYPERLINK',),
    0x0168: ('PHONETIC', 1),
    0x0169: ('AVERAGEA',),
    0x016A: ('MAXA',),
    0x016B: ('MINA',),
    0x016C: ('STDEVPA',),
    0x016D: ('VARPA',),
    0x016E: ('STDEVA',),
    0x016F: ('VARA',),
    0x0170: ('BAHTTEXT', 1),
    0x0171: ('THAIDAYOFWEEK', 1),
    0x0172: ('THAIDIGIT', 1),
    0x0173: ('THAIMONTHOFYEAR', 1),
    0x0174: ('THAINUMSOUND', 1),
    0x0175: ('THAINUMSTRING', 1),
    0x0176: ('THAISTRINGLENGTH', 1),
    0x0177: ('ISTHAIDIGIT', 1),
    0x0178: ('ROUNDBAHTDOWN', 1),
    0x0179: ('ROUNDBAHTUP', 1),
    0x017A: ('THAIYEAR', 1),
    0x017B: ('RTD',),
    # From https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/90a52fcb-ce63-497f-a3d3-173c42d82242
    0x017C: ('CUBEVALUE',),
    0x017D: ('CUBEMEMBER',),
    0x017E: ('CUBEMEMBERPROPERTY', 3),
    0x017F: ('CUBERANKEDMEMBER',),
    0x0180: ('HEX2BIN',),
    0x0181: ('HEX2DEC', 1),
    0x0182: ('HEX2OCT',),
    0x0183: ('DEC2BIN',),
    0x0184: ('DEC2HEX',),
    0x0185: ('DEC2OCT',),
    0x0186: ('OCT2BIN',),
    0x0187: ('OCT2HEX',),
    0x0188: ('OCT2DEC', 1),
    0x0189: ('BIN2DEC', 1),
    0x018A: ('BIN2OCT',),
    0x018B: ('BIN2HEX',),
    0x018C: ('IMSUB', 2),
    0x018D: ('IMDIV', 2),
    0x018E: ('IMPOWER', 2),
    0x018F: ('IMABS', 1),
    0x0190: ('IMSQRT', 1),
    0x0191: ('IMLN', 1),
    0x0192: ('IMLOG2', 1),
    0x0193: ('IMLOG10', 1),
    0x0194: ('IMSIN', 1),
    0x0195: ('IMCOS', 1),
    0x0196: ('IMEXP', 1),
    0x0197: ('IMARGUMENT', 1),
    0x0198: ('IMCONJUGATE', 1),
    0x0199: ('IMAGINARY', 1),
    0x019A: ('IMREAL', 1),
    0x019B: ('COMPLEX',),
    0x019C: ('IMSUM',),
    0x019D: ('IMPRODUCT',),
    0x019E: ('SERIESSUM', 4),
    0x019F: ('FACTDOUBLE', 1),
    0x01A0: ('SQRTPI', 1),
    0x01A1: ('QUOTIENT', 2),
    0x01A2: ('DELTA',),
    0x01A3: ('GESTEP',),
    0x01A4: ('ISEVEN', 1),
    0x01A5: ('ISODD', 1),
    0x01A6: ('MROUND', 2),
    0x01A7: ('ERF',),
    0x01A8: ('ERFC', 1),
    0x01A9: ('BESSELJ', 2),
    0x01AA: ('BESSELK', 2),
    0x01AB: ('BESSELY', 2),
    0x01AC: ('BESSELI', 2),
    0x01AD: ('XIRR',),
    0x01AE: ('XNPV', 3),
    0x01AF: ('PRICEMAT',),
    0x01B0: ('YIELDMAT',),
    0x01B1: ('INTRATE',),
    0x01B2: ('RECEIVED',),
    0x01B3: ('DISC',),
    0x01B4: ('PRICEDISC',),
    0x01B5: ('YIELDDISC',),
    0x01B6: ('TBILLEQ', 3),
    0x01B7: ('TBILLPRICE', 3),
    0x01B8: ('TBILLYIELD', 3),
    0x01B9: ('PRICE',),
    0x01BA: ('YIELD',),
    0x01BB: ('DOLLARDE', 2),
    0x01BC: ('DOLLARFR', 2),
    0x01BD: ('NOMINAL', 2),
    0x01BE: ('EFFECT', 2),
    0x01BF: ('CUMPRINC', 6),
    0x01C0: ('CUMIPMT', 6),
    0x01C1: ('EDATE', 2),
    0x01C2: ('EOMONTH', 2),
    0x01C3: ('YEARFRAC',),
    0x01C4: ('COUPDAYBS',),
    0x01C5: ('COUPDAYS',),
    0x01C6: ('COUPDAYSNC',),
    0x01C7: ('COUPNCD',),
    0x01C8: ('COUPNUM',),
    0x01C9: ('COUPPCD',),
    0x01CA: ('DURATION',),
    0x01CB: ('MDURATION',),
    0x01CC: ('ODDLPRICE',),
    0x01CD: ('ODDLYIELD',),
    0x01CE: ('ODDFPRICE',),
    0x01CF: ('ODDFYIELD',),
    0x01D0: ('RANDBETWEEN', 2),
    0x01D1: ('WEEKNUM',),
    0x01D2: ('AMORDEGRC',),
    0x01D3: ('AMORLINC',),
    0x01D5: ('ACCRINT',),
    0x01D6: ('ACCRINTM',),
    0x01D7: ('WORKDAY',),
    0x01D8: ('NETWORKDAYS',),
    0x01D9: ('GCD',),
    0x01DA: ('MULTINOMIAL',),
    0x01DB: ('LCM',),
    0x01DC: ('FVSCHEDULE', 2),
    0x01DD: ('CUBEKPIMEMBER',),
    0x01DE: ('CUBESET',),
    0x01DF: ('CUBESETCOUNT', 1),
    0x01E0: ('IFERROR', 2),
    0x01E1: ('COUNTIFS',),
    0x01E2: ('SUMIFS',),
    0x01E3: ('AVERAGEIF',),

    # https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/0b8acba5-86d2-4854-836e-0afaee743d44
    0x8000: ('BEEP',),
    0x8001: ('OPEN',),
    0x8002: ('OPEN.LINKS',),
    0x8003: ('CLOSE.ALL',),
    0x8004: ('SAVE',),
    0x8005: ('SAVE.AS',),
    0x8006: ('FILE.DELETE',),
    0x8007: ('PAGE.SETUP',),
    0x8008: ('PRINT',),
    0x8009: ('PRINTER.SETUP',),
    0x800A: ('QUIT',),
    0x800B: ('NEW.WINDOW',),
    0x800C: ('ARRANGE.ALL',),
    0x800D: ('WINDOW.SIZE',),
    0x800E: ('WINDOW.MOVE',),
    0x800F: ('FULL',),
    0x8010: ('CLOSE',),
    0x8011: ('RUN',),
    0x8016: ('SET.PRINT.AREA',),
    0x8017: ('SET.PRINT.TITLES',),
    0x8018: ('SET.PAGE.BREAK',),
    0x8019: ('REMOVE.PAGE.BREAK',),
    0x801A: ('FONT',),
    0x801B: ('DISPLAY',),
    0x801C: ('PROTECT.DOCUMENT',),
    0x801D: ('PRECISION',),
    0x801E: ('A1.R1C1',),
    0x801F: ('CALCULATE.NOW',),
    0x8020: ('CALCULATION',),
    0x8022: ('DATA.FIND',),
    0x8023: ('EXTRACT',),
    0x8024: ('DATA.DELETE',),
    0x8025: ('SET.DATABASE',),
    0x8026: ('SET.CRITERIA',),
    0x8027: ('SORT',),
    0x8028: ('DATA.SERIES',),
    0x8029: ('TABLE',),
    0x802A: ('FORMAT.NUMBER',),
    0x802B: ('ALIGNMENT',),
    0x802C: ('STYLE',),
    0x802D: ('BORDER',),
    0x802E: ('CELL.PROTECTION',),
    0x802F: ('COLUMN.WIDTH',),
    0x8030: ('UNDO',),
    0x8031: ('CUT',),
    0x8032: ('COPY',),
    0x8033: ('PASTE',),
    0x8034: ('CLEAR',),
    0x8035: ('PASTE.SPECIAL',),
    0x8036: ('EDIT.DELETE',),
    0x8037: ('INSERT',),
    0x8038: ('FILL.RIGHT',),
    0x8039: ('FILL.DOWN',),
    0x803D: ('DEFINE.NAME',),
    0x803E: ('CREATE.NAMES',),
    0x803F: ('FORMULA.GOTO',),
    0x8040: ('FORMULA.FIND',),
    0x8041: ('SELECT.LAST.CELL',),
    0x8042: ('SHOW.ACTIVE.CELL',),
    0x8043: ('GALLERY.AREA',),
    0x8044: ('GALLERY.BAR',),
    0x8045: ('GALLERY.COLUMN',),
    0x8046: ('GALLERY.LINE',),
    0x8047: ('GALLERY.PIE',),
    0x8048: ('GALLERY.SCATTER',),
    0x8049: ('COMBINATION',),
    0x804A: ('PREFERRED',),
    0x804B: ('ADD.OVERLAY',),
    0x804C: ('GRIDLINES',),
    0x804D: ('SET.PREFERRED',),
    0x804E: ('AXES',),
    0x804F: ('LEGEND',),
    0x8050: ('ATTACH.TEXT',),
    0x8051: ('ADD.ARROW',),
    0x8052: ('SELECT.CHART',),
    0x8053: ('SELECT.PLOT.AREA',),
    0x8054: ('PATTERNS',),
    0x8055: ('MAIN.CHART',),
    0x8056: ('OVERLAY',),
    0x8057: ('SCALE',),
    0x8058: ('FORMAT.LEGEND',),
    0x8059: ('FORMAT.TEXT',),
    0x805A: ('EDIT.REPEAT',),
    0x805B: ('PARSE',),
    0x805C: ('JUSTIFY',),
    0x805D: ('HIDE',),
    0x805E: ('UNHIDE',),
    0x805F: ('WORKSPACE',),
    0x8060: ('FORMULA',),
    0x8061: ('FORMULA.FILL',),
    0x8062: ('FORMULA.ARRAY',),
    0x8063: ('DATA.FIND.NEXT',),
    0x8064: ('DATA.FIND.PREV',),
    0x8065: ('FORMULA.FIND.NEXT',),
    0x8066: ('FORMULA.FIND.PREV',),
    0x8067: ('ACTIVATE',),
    0x8068: ('ACTIVATE.NEXT',),
    0x8069: ('ACTIVATE.PREV',),
    0x806A: ('UNLOCKED.NEXT',),
    0x806B: ('UNLOCKED.PREV',),
    0x806C: ('COPY.PICTURE',),
    0x806D: ('SELECT',),
    0x806E: ('DELETE.NAME',),
    0x806F: ('DELETE.FORMAT',),
    0x8070: ('VLINE',),
    0x8071: ('HLINE',),
    0x8072: ('VPAGE',),
    0x8073: ('HPAGE',),
    0x8074: ('VSCROLL',),
    0x8075: ('HSCROLL',),
    0x8076: ('ALERT',),
    0x8077: ('NEW',),
    0x8078: ('CANCEL.COPY',),
    0x8079: ('SHOW.CLIPBOARD',),
    0x807A: ('MESSAGE',),
    0x807C: ('PASTE.LINK',),
    0x807D: ('APP.ACTIVATE',),
    0x807E: ('DELETE.ARROW',),
    0x807F: ('ROW.HEIGHT',),
    0x8080: ('FORMAT.MOVE',),
    0x8081: ('FORMAT.SIZE',),
    0x8082: ('FORMULA.REPLACE',),
    0x8083: ('SEND.KEYS',),
    0x8084: ('SELECT.SPECIAL',),
    0x8085: ('APPLY.NAMES',),
    0x8086: ('REPLACE.FONT',),
    0x8087: ('FREEZE.PANES',),
    0x8088: ('SHOW.INFO',),
    0x8089: ('SPLIT',),
    0x808A: ('ON.WINDOW',),
    0x808B: ('ON.DATA',),
    0x808C: ('DISABLE.INPUT',),
    0x808E: ('OUTLINE',),
    0x808F: ('LIST.NAMES',),
    0x8090: ('FILE.CLOSE',),
    0x8091: ('SAVE.WORKBOOK',),
    0x8092: ('DATA.FORM',),
    0x8093: ('COPY.CHART',),
    0x8094: ('ON.TIME',),
    0x8095: ('WAIT',),
    0x8096: ('FORMAT.FONT',),
    0x8097: ('FILL.UP',),
    0x8098: ('FILL.LEFT',),
    0x8099: ('DELETE.OVERLAY',),
    0x809B: ('SHORT.MENUS',),
    0x809F: ('SET.UPDATE.STATUS',),
    0x80A1: ('COLOR.PALETTE',),
    0x80A2: ('DELETE.STYLE',),
    0x80A3: ('WINDOW.RESTORE',),
    0x80A4: ('WINDOW.MAXIMIZE',),
    0x80A6: ('CHANGE.LINK',),
    0x80A7: ('CALCULATE.DOCUMENT',),
    0x80A8: ('ON.KEY',),
    0x80A9: ('APP.RESTORE',),
    0x80AA: ('APP.MOVE',),
    0x80AB: ('APP.SIZE',),
    0x80AC: ('APP.MINIMIZE',),
    0x80AD: ('APP.MAXIMIZE',),
    0x80AE: ('BRING.TO.FRONT',),
    0x80AF: ('SEND.TO.BACK',),
    0x80B9: ('MAIN.CHART.TYPE',),
    0x80BA: ('OVERLAY.CHART.TYPE',),
    0x80BB: ('SELECT.END',),
    0x80BC: ('OPEN.MAIL',),
    0x80BD: ('SEND.MAIL',),
    0x80BE: ('STANDARD.FONT',),
    0x80BF: ('CONSOLIDATE',),
    0x80C0: ('SORT.SPECIAL',),
    0x80C1: ('GALLERY.3D.AREA',),
    0x80C2: ('GALLERY.3D.COLUMN',),
    0x80C3: ('GALLERY.3D.LINE',),
    0x80C4: ('GALLERY.3D.PIE',),
    0x80C5: ('VIEW.3D',),
    0x80C6: ('GOAL.SEEK',),
    0x80C7: ('WORKGROUP',),
    0x80C8: ('FILL.GROUP',),
    0x80C9: ('UPDATE.LINK',),
    0x80CA: ('PROMOTE',),
    0x80CB: ('DEMOTE',),
    0x80CC: ('SHOW.DETAIL',),
    0x80CE: ('UNGROUP',),
    0x80CF: ('OBJECT.PROPERTIES',),
    0x80D0: ('SAVE.NEW.OBJECT',),
    0x80D1: ('SHARE',),
    0x80D2: ('SHARE.NAME',),
    0x80D3: ('DUPLICATE',),
    0x80D4: ('APPLY.STYLE',),
    0x80D5: ('ASSIGN.TO.OBJECT',),
    0x80D6: ('OBJECT.PROTECTION',),
    0x80D7: ('HIDE.OBJECT',),
    0x80D8: ('SET.EXTRACT',),
    0x80D9: ('CREATE.PUBLISHER',),
    0x80DA: ('SUBSCRIBE.TO',),
    0x80DB: ('ATTRIBUTES',),
    0x80DC: ('SHOW.TOOLBAR',),
    0x80DE: ('PRINT.PREVIEW',),
    0x80DF: ('EDIT.COLOR',),
    0x80E0: ('SHOW.LEVELS',),
    0x80E1: ('FORMAT.MAIN',),
    0x80E2: ('FORMAT.OVERLAY',),
    0x80E3: ('ON.RECALC',),
    0x80E4: ('EDIT.SERIES',),
    0x80E5: ('DEFINE.STYLE',),
    0x80F0: ('LINE.PRINT',),
    0x80F3: ('ENTER.DATA',),
    0x80F9: ('GALLERY.RADAR',),
    0x80FA: ('MERGE.STYLES',),
    0x80FB: ('EDITION.OPTIONS',),
    0x80FC: ('PASTE.PICTURE',),
    0x80FD: ('PASTE.PICTURE.LINK',),
    0x80FE: ('SPELLING',),
    0x8100: ('ZOOM',),
    0x8103: ('INSERT.OBJECT',),
    0x8104: ('WINDOW.MINIMIZE',),
    0x8109: ('SOUND.NOTE',),
    0x810A: ('SOUND.PLAY',),
    0x810B: ('FORMAT.SHAPE',),
    0x810C: ('EXTEND.POLYGON',),
    0x810D: ('FORMAT.AUTO',),
    0x8110: ('GALLERY.3D.BAR',),
    0x8111: ('GALLERY.3D.SURFACE',),
    0x8112: ('FILL.AUTO',),
    0x8114: ('CUSTOMIZE.TOOLBAR',),
    0x8115: ('ADD.TOOL',),
    0x8116: ('EDIT.OBJECT',),
    0x8117: ('ON.DOUBLECLICK',),
    0x8118: ('ON.ENTRY',),
    0x8119: ('WORKBOOK.ADD',),
    0x811A: ('WORKBOOK.MOVE',),
    0x811B: ('WORKBOOK.COPY',),
    0x811C: ('WORKBOOK.OPTIONS',),
    0x811D: ('SAVE.WORKSPACE',),
    0x8120: ('CHART.WIZARD',),
    0x8121: ('DELETE.TOOL',),
    0x8122: ('MOVE.TOOL',),
    0x8123: ('WORKBOOK.SELECT',),
    0x8124: ('WORKBOOK.ACTIVATE',),
    0x8125: ('ASSIGN.TO.TOOL',),
    0x8127: ('COPY.TOOL',),
    0x8128: ('RESET.TOOL',),
    0x8129: ('CONSTRAIN.NUMERIC',),
    0x812A: ('PASTE.TOOL',),
    0x812E: ('WORKBOOK.NEW',),
    0x8131: ('SCENARIO.CELLS',),
    0x8132: ('SCENARIO.DELETE',),
    0x8133: ('SCENARIO.ADD',),
    0x8134: ('SCENARIO.EDIT',),
    0x8135: ('SCENARIO.SHOW',),
    0x8136: ('SCENARIO.SHOW.NEXT',),
    0x8137: ('SCENARIO.SUMMARY',),
    0x8138: ('PIVOT.TABLE.WIZARD',),
    0x8139: ('PIVOT.FIELD.PROPERTIES',),
    0x813A: ('PIVOT.FIELD',),
    0x813B: ('PIVOT.ITEM',),
    0x813C: ('PIVOT.ADD.FIELDS',),
    0x813E: ('OPTIONS.CALCULATION',),
    0x813F: ('OPTIONS.EDIT',),
    0x8140: ('OPTIONS.VIEW',),
    0x8141: ('ADDIN.MANAGER',),
    0x8142: ('MENU.EDITOR',),
    0x8143: ('ATTACH.TOOLBARS',),
    0x8144: ('VBAActivate',),
    0x8145: ('OPTIONS.CHART',),
    0x8148: ('VBA.INSERT.FILE',),
    0x814A: ('VBA.PROCEDURE.DEFINITION',),
    0x8150: ('ROUTING.SLIP',),
    0x8152: ('ROUTE.DOCUMENT',),
    0x8153: ('MAIL.LOGON',),
    0x8156: ('INSERT.PICTURE',),
    0x8157: ('EDIT.TOOL',),
    0x8158: ('GALLERY.DOUGHNUT',),
    0x815E: ('CHART.TREND',),
    0x8160: ('PIVOT.ITEM.PROPERTIES',),
    0x8162: ('WORKBOOK.INSERT',),
    0x8163: ('OPTIONS.TRANSITION',),
    0x8164: ('OPTIONS.GENERAL',),
    0x8172: ('FILTER.ADVANCED',),
    0x8175: ('MAIL.ADD.MAILER',),
    0x8176: ('MAIL.DELETE.MAILER',),
    0x8177: ('MAIL.REPLY',),
    0x8178: ('MAIL.REPLY.ALL',),
    0x8179: ('MAIL.FORWARD',),
    0x817A: ('MAIL.NEXT.LETTER',),
    0x817B: ('DATA.LABEL',),
    0x817C: ('INSERT.TITLE',),
    0x817D: ('FONT.PROPERTIES',),
    0x817E: ('MACRO.OPTIONS',),
    0x817F: ('WORKBOOK.HIDE',),
    0x8180: ('WORKBOOK.UNHIDE',),
    0x8181: ('WORKBOOK.DELETE',),
    0x8182: ('WORKBOOK.NAME',),
    0x8184: ('GALLERY.CUSTOM',),
    0x8186: ('ADD.CHART.AUTOFORMAT',),
    0x8187: ('DELETE.CHART.AUTOFORMAT',),
    0x8188: ('CHART.ADD.DATA',),
    0x8189: ('AUTO.OUTLINE',),
    0x818A: ('TAB.ORDER',),
    0x818B: ('SHOW.DIALOG',),
    0x818C: ('SELECT.ALL',),
    0x818D: ('UNGROUP.SHEETS',),
    0x818E: ('SUBTOTAL.CREATE',),
    0x818F: ('SUBTOTAL.REMOVE',),
    0x8190: ('RENAME.OBJECT',),
    0x819C: ('WORKBOOK.SCROLL',),
    0x819D: ('WORKBOOK.NEXT',),
    0x819E: ('WORKBOOK.PREV',),
    0x819F: ('WORKBOOK.TAB.SPLIT',),
    0x81A0: ('FULL.SCREEN',),
    0x81A1: ('WORKBOOK.PROTECT',),
    0x81A4: ('SCROLLBAR.PROPERTIES',),
    0x81A5: ('PIVOT.SHOW.PAGES',),
    0x81A6: ('TEXT.TO.COLUMNS',),
    0x81A7: ('FORMAT.CHARTTYPE',),
    0x81A8: ('LINK.FORMAT',),
    0x81A9: ('TRACER.DISPLAY',),
    0x81AE: ('TRACER.NAVIGATE',),
    0x81AF: ('TRACER.CLEAR',),
    0x81B0: ('TRACER.ERROR',),
    0x81B1: ('PIVOT.FIELD.GROUP',),
    0x81B2: ('PIVOT.FIELD.UNGROUP',),
    0x81B3: ('CHECKBOX.PROPERTIES',),
    0x81B4: ('LABEL.PROPERTIES',),
    0x81B5: ('LISTBOX.PROPERTIES',),
    0x81B6: ('EDITBOX.PROPERTIES',),
    0x81B7: ('PIVOT.REFRESH',),
    0x81B8: ('LINK.COMBO',),
    0x81B9: ('OPEN.TEXT',),
    0x81BA: ('HIDE.DIALOG',),
    0x81BB: ('SET.DIALOG.FOCUS',),
    0x81BC: ('ENABLE.OBJECT',),
    0x81BD: ('PUSHBUTTON.PROPERTIES',),
    0x81BE: ('SET.DIALOG.DEFAULT',),
    0x81BF: ('FILTER',),
    0x81C0: ('FILTER.SHOW.ALL',),
    0x81C1: ('CLEAR.OUTLINE',),
    0x81C2: ('FUNCTION.WIZARD',),
    0x81C3: ('ADD.LIST.ITEM',),
    0x81C4: ('SET.LIST.ITEM',),
    0x81C5: ('REMOVE.LIST.ITEM',),
    0x81C6: ('SELECT.LIST.ITEM',),
    0x81C7: ('SET.CONTROL.VALUE',),
    0x81C8: ('SAVE.COPY.AS',),
    0x81CA: ('OPTIONS.LISTS.ADD',),
    0x81CB: ('OPTIONS.LISTS.DELETE',),
    0x81CC: ('SERIES.AXES',),
    0x81CD: ('SERIES.X',),
    0x81CE: ('SERIES.Y',),
    0x81CF: ('ERRORBAR.X',),
    0x81D0: ('ERRORBAR.Y',),
    0x81D1: ('FORMAT.CHART',),
    0x81D2: ('SERIES.ORDER',),
    0x81D3: ('MAIL.LOGOFF',),
    0x81D4: ('CLEAR.ROUTING.SLIP',),
    0x81D5: ('APP.ACTIVATE.MICROSOFT',),
    0x81D6: ('MAIL.EDIT.MAILER',),
    0x81D7: ('ON.SHEET',),
    0x81D8: ('STANDARD.WIDTH',),
    0x81D9: ('SCENARIO.MERGE',),
    0x81DA: ('SUMMARY.INFO',),
    0x81DB: ('FIND.FILE',),
    0x81DC: ('ACTIVE.CELL.FONT',),
    0x81DD: ('ENABLE.TIPWIZARD',),
    0x81DE: ('VBA.MAKE.ADDIN',),
    0x81E0: ('INSERTDATATABLE',),
    0x81E1: ('WORKGROUP.OPTIONS',),
    0x81E2: ('MAIL.SEND.MAILER',),
    0x81E5: ('AUTOCORRECT',),
    0x81E9: ('POST.DOCUMENT',),
    0x81EB: ('PICKLIST',),
    0x81ED: ('VIEW.SHOW',),
    0x81EE: ('VIEW.DEFINE',),
    0x81EF: ('VIEW.DELETE',),
    0x81FD: ('SHEET.BACKGROUND',),
    0x81FE: ('INSERT.MAP.OBJECT',),
    0x81FF: ('OPTIONS.MENONO',),
    0x8205: ('MSOCHECKS',),
    0x8206: ('NORMAL',),
    0x8207: ('LAYOUT',),
    0x8208: ('RM.PRINT.AREA',),
    0x8209: ('CLEAR.PRINT.AREA',),
    0x820A: ('ADD.PRINT.AREA',),
    0x820B: ('MOVE.BRK',),
    0x8221: ('HIDECURR.NOTE',),
    0x8222: ('HIDEALL.NOTES',),
    0x8223: ('DELETE.NOTE',),
    0x8224: ('TRAVERSE.NOTES',),
    0x8225: ('ACTIVATE.NOTES',),
    0x826C: ('PROTECT.REVISIONS',),
    0x826D: ('UNPROTECT.REVISIONS',),
    0x8287: ('OPTIONS.ME',),
    0x828D: ('WEB.PUBLISH',),
    0x829B: ('NEWWEBQUERY',),
    0x82A1: ('PIVOT.TABLE.CHART',),
    0x82F1: ('OPTIONS.SAVE',),
    0x82F3: ('OPTIONS.SPELL',),
    0x8328: ('HIDEALL.INKANNOTS',),
}
