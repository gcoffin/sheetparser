Documents and backends
======================

We load first a backend - that's the module that will read the Excel
sheet and provide the information to the library. There are 5 provided
backends: 

 * one based on `xlrd`_, that can read xls file and xlsx without
formatting,
 * one based on `openpyxl`_ that can read xlsx files with some
formatting information 
 * one is based on win32com and the actual Excel program, with serious performance issues
 * raw provides an interface for data stored as list, as well as csv files
 * `pdfminer`_ provides an interface for pdf files. This feature is experimental and is limited by the amount of information that pdf files can provide.


.. _xlrd: https://pypi.python.org/pypi/xlrd

.. _openpyxl: https://pypi.python.org/pypi/openpyxl

.. _pdfminer: https://pypi.python.org/pypi/pdfminer.six
