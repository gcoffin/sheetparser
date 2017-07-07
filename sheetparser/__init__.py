"""
Sheet Parser module for Python
==============================

sheetparse in a Python module developed to simplify data sheet
loading, specifically Excel documents created by humans with many
tables and complex layout.

What is data sheet parsing?

People use data sheets (Excel, OpenOffice, etc.) to store and exchange
data, and present their data in a way that pleases their eyes. Reading
those sheets from a program can be quite painful. Usual API like
simple workbooks with one table by sheet and the header on the first
row. This is generally not the case with sheets created by humans.

sheetparser can read from data sheets (Excel reading is provided
through packages such as openpyxl and xlrd), and configured to recognize
advanced layouts.

(c) 2017 Guillaume Coffin

Licensed under the GNU General Public Licence v3 (GPLv3)
https://www.gnu.org/licenses/gpl-3.0.txt
"""

__version__ = "0.1.1"
__author__ = "Guillaume Coffin <guill.coffin@gmail.com>"
__license__ = 'GPL v3'
__copyright__ = 'Copyright 2017 Guillaume Coffin'

from .utils import *
from .documents import *
from .patterns import *
from .results import *

__all__ = ['DoesntMatchException', 'QuickPrint', 'Sequence', 'Many', 'Maybe', 'Workbook', 'Range',
           'Sheet', 'no_vertical', 'no_horizontal', 'empty_line',
           'TableTransform', 'TableNotEmpty', 'FillData',
           'get_value', 'Match', 'StripLine', 'GetValue', 'match_if','KeepOnly',
           'IgnoreIf',
           'HeaderTableTransform', 'RepeatExisting', 'RemoveEmptyLines', 'ToMap',
           'MergeHeader', 'Transpose', 'ToDate', 'Table', 'DEFAULT_TRANSFORMS',
           'CellRange', 'OrPattern',
           'FlexibleRange', 'Line', 'Empty', 'Rows','VisibleRows',
           'Columns', 'Document', 'BORDER_TOP', 'BORDER_LEFT', 'BORDER_BOTTOM',
           'BORDER_RIGHT', 'BORDERS_VERTICAL', 'BORDERS_HORIZONTAL', 'EMPTY_CELL',
           'numrow', 'RbColIterator', 'RbRowIterator',
           'PythonObjectContext', 'ResultContext', 'ListContext', 'DebugContext',
           'load_backend', 'load_workbook']
