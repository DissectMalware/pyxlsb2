pyxlsb2
======

``pyxlsb2`` (a variant of pyxlsb) is an Excel 2007-2010 Binary Workbook (xlsb) parser for Python.

``pyxslb2`` offers the following improvements/changes in comparison to pyxlsb:

1. By default, keeps all data in memory instead of creating temporary files. This is mainly to speed up the processing and also not changing the local filesystem during the processing.
2. relies on both "xl\\workbook.bin" and "xl\\_rels\\workbook.bin.rels" to load locate boundsheets. As a result, it can load all worksheets as well as all macrosheets.
3. extracts macro formulas:
 * accurately shows the formulas
 * supports A1 addressing
 * supports external addressing (partially implemented))


Install
-----

1. Installing the whl file

 Download \.whl file from release section
 
     pip install -U [path to whl file]
    
2. Installing the latest development 

 Download the latest version
 
     wget https://github.com/DissectMalware/pyxlsb2/archive/master.zip

 Extract the zip file and go to the extracted directory
 
     python setup.py install --user


Usage
-----

The module exposes an ``open_workbook(name)`` method (similar to Xlrd and
OpenPyXl) for opening XLSB files. The Workbook object representing the file is
returned.

.. code:: python

   from pyxlsb2 import open_workbook
   with open_workbook('Book1.xlsb') as wb:
       # Do stuff with wb

The Workbook object exposes a ``get_sheet_by_index(idx)`` and
``get_sheet_by_name(name)`` method to retrieve Worksheet instances.

.. code:: python

   # Using the sheet index (0-based, unlike VBA)
   with wb.get_sheet_by_index(0) as sheet:
       # Do stuff with sheet

   # Using the sheet name
   with wb.get_sheet_by_name('Sheet1') as sheet:
       # Do stuff with sheet

A ``sheets`` property containing the sheet names is available on the Workbook
instance.

The ``rows()`` method will hand out an iterator to read the worksheet rows. The
Worksheet object is also directly iterable and is equivalent to calling
``rows()``.

.. code:: python

   # You can use .rows(sparse=False) to include empty rows
   for row in sheet.rows():
       print(row)
   # [Cell(r=0, c=0, v='TEXT'), Cell(r=0, c=1, v=42.1337)]

*NOTE*: Iterating the same Worksheet instance multiple times in parallel (nested
``for`` for instance) will yield unexpected results, retrieve more instances
instead.

Note that dates will appear as floats. You must use the ``convert_date(date)``
method from the corresponding Workbook instance to turn them into ``datetime``.

.. code:: python

   print(wb.convert_date(41235.45578))
   # datetime.datetime(2012, 11, 22, 10, 56, 19)


Example
-------

Converting a workbook to CSV:

.. code:: python

   import csv
   from pyxlsb2 import open_workbook

   with open_workbook('Book1.xlsb') as wb:
       for name in wb.sheets:
           with wb.get_sheet_by_name(name) as sheet:
               with open(name + '.csv', 'w') as f:
                   writer = csv.writer(f)
                   for row in sheet.rows():
                       writer.writerow([c.v for c in row])

Limitations 
-----------

Non exhaustive list of things that are currently not supported:

-  Style and formatting *WIP*
-  Rich text cells (formatting is lost, but getting the text works)
-  Encrypted (password protected) workbooks
-  Comments and other annotations
-  Writing (out of scope)


