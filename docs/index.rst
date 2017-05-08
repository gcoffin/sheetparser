.. sheetparser documentation master file, created by
   sphinx-quickstart on Fri May 05 21:59:06 2017.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

sheetparser: the Excel scraper
==============================

sheetparser is a library for extracting information from Excel sheets
that contain complex or variable layouts of tables.

Obtaining data from various sources can be very painful, and loading
Excel sheets that were designed by humans for humans is especially
difficult. The focus of the persons who create those sheet is first to
display the information in a way that pleases their eyes or can
convince others, and readability by a computer is low on the list of
priorities. Also as time goes, they add intermediate lines or columns,
or add more information. The systematic loading of historical
information can then become a very heavy task.

The purpose of this package is to simplify the data extraction of
those tables. Complex and flexible layouts can be implemented in a few
lines.

Contents:
=========

.. toctree::
   :maxdepth: 2

   introduction
   patterns
   transformations
   results
   backends
   
Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`
