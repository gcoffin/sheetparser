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
Licensed under the GNU Lesser General Public Licence v3 (LGPLv3)
https://www.gnu.org/copyleft/lesser.html
"""

__version__ = "0.1a1"
__author__ = "Guillaume Coffin <guill.coffin@gmail.com>"



from .utils import *
from .documents import *
from .patterns import *
from .results import *

__all__ = ['DoesntMatchException', 'QuickPrint', 'PythonObjectContext',
           'ListContext', 'Sequence', 'Many', 'Maybe', 'Workbook', 'Range',
           'Sheet', 'no_vertical', 'no_horizontal', 'empty_line',
           'TableNotEmpty', 'FillData',
           'get_value', 'Match', 'StripLine','GetValue',
           'HeaderTableTransform', 'RepeatExisting', 'RemoveEmptyLines', 'ToMap',
           'MergeHeader', 'Transpose', 'ToDate', 'Table', 'DEFAULT_TRANFORMS',
           'CellRange', 'ResultContext', 'FlexibleRange', 'Line', 'Empty', 'Rows',
           'Columns', 'Document', 'BORDER_TOP', 'BORDER_LEFT', 'BORDER_BOTTOM',
           'BORDER_RIGHT', 'BORDERS_VERTICAL', 'BORDERS_HORIZONTAL', 'EMPTY_CELL',
           'numrow', 'RbColIterator', 'RbRowIterator', 'load_backend', 'load_workbook']

