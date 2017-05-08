Documents and backends
======================

We load first a backend - that's the module that will read the Excel
sheet and provide the information to the library. There are 3 provided
backends: one based on `xlrd`_, that can read xls file and xlsx without
formatting, one based on `openpyxl`_ that can read xlsx files with some
formatting information and the last one is based on win32com and the
actual Excel program, with serious performance issues.  


.. _xlrd: https://pypi.python.org/pypi/xlrd

.. _openpyxl: https://pypi.python.org/pypi/openpyxl
