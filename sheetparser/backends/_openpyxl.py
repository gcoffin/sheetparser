"""uses openpyxl as the engine to read the Excel sheet."""

from __future__ import print_function
import os
import six
import logging

import openpyxl
import openpyxl.styles

from ..documents import (BORDER_TOP, BORDER_LEFT, 
                         BORDER_BOTTOM, BORDER_RIGHT, 
                         CellRange, SheetDocument, WorkbookDocument,
                         load_workbook)
from ..utils import EMPTY_CELL


SHEETSTATE_VISIBLE = openpyxl.worksheet.Worksheet.SHEETSTATE_VISIBLE
logger = logging.getLogger('sheetparser')


class EmptyCell(object):
    value = EMPTY_CELL

    def is_empty(self):
        return True

    def has_borders(self, mask):
        return False


XL_EMPTY_CELL = EmptyCell()


class opxlCell(object):
    def __init__(self, value, cell, wksheet_fmt, is_merged):
        self._cell = cell
        self.value = EMPTY_CELL if value is None else value
        self._wksheet_fmt = wksheet_fmt
        self._border_mask = None
        self.is_merged = is_merged

    @property
    def border_mask(self):
        if self._border_mask is None:
            border = self._cell.border
            if border is None:
                self._border_mask = 0
            else:
                self._border_mask = (
                    (BORDER_TOP*(border.top.style is not None)) |
                    (BORDER_LEFT*(border.left.style is not None)) |
                    (BORDER_BOTTOM*(border.bottom.style is not None)) |
                    (BORDER_RIGHT*(border.right.style is not None)))
        return self._border_mask

    def has_borders(self, mask):
        return bool(self.border_mask & mask)

    @property
    def is_filled(self):
        return self._cell.fill.patternType == 'solid'

    @property
    def color(self):
        return {k:float(v) for k,v in self._cell.fill.fgColor}

    def color_hash(self):
        # reading a color from openpyxl is feasible only when
        # we're lucky (color is properly indexed)
        # so we focus on the very basic: color or not, and some hash
        try:
            return self._cell.fill.start_color._key
        except:
            return None
    
    def is_empty(self):
        return self.value == EMPTY_CELL


class opxlExcelSheet(CellRange, SheetDocument):
    def __init__(self, wksheet_data, wksheet_fmt=None):
        self.name = wksheet_data.title
        self.wksheet_data = wksheet_data
        self.wksheet_fmt = wksheet_fmt
        self.merged = {}
        if self.wksheet_fmt:
            for crange in self.wksheet_fmt.merged_cell_ranges:
                clo, rlo, chi, rhi = openpyxl.utils.range_boundaries(crange)
                for rowx in range(rlo, rhi):
                    for colx in range(clo, chi):
                        if (rlo, clo) != (rowx, colx):
                            self.merged[rowx, colx] = (rlo, clo)
        self.top, self.left = 1, 1  # wksheet.min_row, wksheet.min_column
        self.bottom = wksheet_data.max_row + 1
        self.right = wksheet_data.max_column + 1

    def is_hidden(self):
        return self.wksheet_data.sheet_state != SHEETSTATE_VISIBLE

    def cell(self, row, col):
        abs_row = self.top + row
        abs_col = self.left + col
        is_merged = False
        if (abs_row, abs_col) in self.merged:
            is_merged = True
            abs_row, abs_col = self.merged[abs_row, abs_col]
        try:
            return opxlCell(
                self.wksheet_data.cell(row=abs_row, column=abs_col).value, 
                (self.wksheet_fmt.cell(row=abs_row, column=abs_col)
                 if self.wksheet_fmt else None),
                self.wksheet_fmt, is_merged)
        except IndexError:
            return XL_EMPTY_CELL

    def __repr__(self):
        return "<opxlExcelSheet %s>" % self.name


class opxlExcelWorkbook(WorkbookDocument):
    """A class to open workbooks and obtain sheets"""
    def __init__(self, filename, with_formatting=True):
        # I'd like to open it readonly but then the merged cells 
        # are not loaded
        if with_formatting:
            self.wbk_fmt = openpyxl.load_workbook(filename=filename)
        else:
            self.wbk_fmt = None
        # and we need the data too because cell values are the formulas!!
        self.wbk_data = openpyxl.load_workbook(filename=filename, data_only=True)
        
    def __iter__(self):
        return (self[s] for s in self.wbk_data.get_sheet_names())

    def __getitem__(self, name_or_id):
        if isinstance(name_or_id, six.string_types):
            return opxlExcelSheet(self.wbk_data.get_sheet_by_name(name_or_id),
                                  self.wbk_fmt.get_sheet_by_name(name_or_id) if self.wbk_fmt else None)
        else:
            return opxlExcelSheet(self.wbk_data.worksheets[id],
                                  self.wbk_fmt.worksheets[id] if self.wbk_fmt else None)

load_workbook = opxlExcelWorkbook
