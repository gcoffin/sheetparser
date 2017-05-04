import struct
import logging

import win32com.client

from ..documents import (BORDER_TOP, BORDER_LEFT, 
                         BORDER_BOTTOM, BORDER_RIGHT, 
                         CellRange, SheetDocument, WorkbookDocument,
                         load_workbook)

from ..utils import EMPTY_CELL

logger = logging.getLogger('sheetparser')

'''works with any Excel version but VERY slow'''

class win32Cell(object):
    BORDER_TOP_ID, BORDER_LEFT_ID, BORDER_BOTTOM_ID, BORDER_RIGHT_ID = 3,1,4,2

    def __init__(self, cell, wksheet):
        self._cell = cell
        self._wksheet = wksheet
        self._border_mask = None
        self.is_merged = cell.MergeCells

    @property
    def border_mask(self):
        if self._border_mask is None:
            borders = self._cell.Borders
            border_mask = 0
            for mask,idx in ((BORDER_TOP, self.BORDER_TOP_ID),
                             (BORDER_LEFT, self.BORDER_LEFT_ID),
                             (BORDER_BOTTOM, self.BORDER_BOTTOM_ID),
                             (BORDER_RIGHT, self.BORDER_RIGHT_ID)):
                border_mask |= mask * (borders[idx].LineStyle != -4142)
            self._border_mask = border_mask
        return self._border_mask

    def has_borders(self, mask):
        return bool(self.border_mask & mask)

    @property
    def color(self):
        b,r,g,a = struct.Struct('4B').unpack(
            struct.Struct('I').pack(
                int(self._cell.Interior.Color)))
        return b,r,g,a

    @property
    def value(self):
        value = self._cell.Value
        if value is None: return EMPTY_CELL
        return value
    
    def is_empty(self):
        return self._cell.Value is None or self._cell.Value == ''


def get_range_boundaries(range):
    top = range.Row
    bottom = top + range.Rows.Count
    left = range.Column
    right = left + range.Columns.Count
    return (top, left, bottom, right)
    

class win32ExcelSheet(CellRange, SheetDocument):
    def __init__(self, wksheet):
        self.name = wksheet.Name
        self.wksheet = wksheet
        self.top, self.left = (1, 1)
        _, _, self.bottom, self.right = get_range_boundaries(wksheet.UsedRange)

    def is_hidden(self):
        return not self.wksheet.Visible

    def cell(self, row, col):
        return win32Cell(self.wksheet.Cells(self.top + row, self.left + col), self)

    def __repr__(self):
        return "<win32ExcelSheet %s>" % self.name


class win32ExcelWorkbook(WorkbookDocument):    
    def __init__(self, filename, with_formatting=True):
        logger.info('Open file %s'%filename)
        self.xc = win32com.client.Dispatch('Excel.Application')
        self.wbk = self.xc.Workbooks.Open(filename, ReadOnly = True, UpdateLinks = False)

    def __iter__(self):
        return (win32ExcelSheet(s) for s in self.wbk.Sheets)

    def __getitem__(self, name_or_id):
        return win32ExcelSheet(self.wbk.Sheets[name_or_id])

load_workbook = win32ExcelWorkbook
