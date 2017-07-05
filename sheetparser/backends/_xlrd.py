"""Provides the interface to open an Excel workbooks, using xlrd as
the engine. Can read xls with formatting, or xlsx without
formatting."""

import xlrd
import six

from ..documents import (BORDER_TOP, BORDER_LEFT, 
                         BORDER_BOTTOM, BORDER_RIGHT, 
                         CellRange, SheetDocument, WorkbookDocument,
                         load_workbook)

class xlrdCell(object):
    def __init__(self, cell, wksheet, is_merged):
        self._cell = cell
        self._wksheet = wksheet
        self._xf_cell = None
        self.is_merged = is_merged

    @property
    def value(self):
        return self._cell.value

    @property
    def _ctype(self):
        return self._cell.ctype

    @property
    def formatting(self):
        if self._xf_cell is None:
            self._xf_cell = XfCell(self._cell.xf_index, self._wksheet)
        return self._xf_cell

    @property
    def is_empty(self):
        return (self._ctype == xlrd.XL_CELL_EMPTY or
                self._ctype == xlrd.XL_CELL_BLANK or
                (self._ctype == xlrd.XL_CELL_TEXT and
                 self.value.strip() == ''))

    def has_borders(self, mask):
        return bool(self.border_mask & mask)

    @property
    def border_mask(self):
        return self.formatting.border_mask

    @property
    def fill(self):
        return self.formatting.fill

    
    @property
    def is_filled(self):
        return self.formatting.is_filled



class Fill(object):
    PATTERN = {0:None,
               1:'solid'}

    def __init__(self,xf_record,book):
        self.type = 'patternFill' # xlrd + formatting --> .xls, only supports patterns
        self.pattern = self.PATTERN[xf_record.background.fill_pattern]
        self.color1 = book.colour_map[xf_record.background.pattern_colour_index]
        self.color2 = book.colour_map[xf_record.background.background_colour_index]

class XfCell(object):
    def __init__(self, xf_index, wksheet):
        self._wksheet = wksheet
        self._xf_record = wksheet.book.xf_list[xf_index]
        self._borders = None
        self._fill = None

    @property
    def border_mask(self):
        if self._borders is None:
            border = self._xf_record.border
            self._borders = ((BORDER_TOP*(border.top_line_style != 0)) |
                            (BORDER_LEFT*(border.left_line_style != 0)) |
                            (BORDER_BOTTOM*(border.bottom_line_style != 0)) |
                            (BORDER_RIGHT*(border.right_line_style != 0)))
        return self._borders

    @property
    def fill(self):
        if self._fill is None:
            self._fill = Fill(self._xf_record,self._wksheet.book)
        return self._fill

    @property
    def is_filled(self):
        return self.fill.type != 'patternFill' or self.fill.pattern is not None


class xlrdExcelSheet(SheetDocument, CellRange):
    def __init__(self, wksheet):
        self.name = wksheet.name
        self.wksheet = wksheet
        self.merged = {}
        for crange in wksheet.merged_cells:
            rlo, rhi, clo, chi = crange
            for rowx in range(rlo, rhi):
                for colx in range(clo, chi):
                    if (rlo,clo) != (rowx,colx):
                        self.merged[rowx, colx] = (rlo, clo)
        self.top, self.left = 0, 0
        self.bottom = wksheet.nrows
        self.right = wksheet.ncols

    def is_hidden(self):
        return self.wksheet.visibility != 0

    def is_hidden_row(self,rowidx):
        row_info = self.wksheet.rowinfo_map.get(rowidx)
        if row_info is None: return False
        return row_info.hidden

    def cell(self, row, col, ignore_merged=False):
        is_merged = False
        if (row,col) in self.merged:
            is_merged = True
            row, col = self.merged[row, col]
        return xlrdCell(self.wksheet.cell(row, col), self.wksheet, is_merged)

    def __repr__(self):
        return "<xlrdExcelSheet %s>" % self.name


class xlrdExcelWorkbook(WorkbookDocument):    
    def __init__(self, filename, with_formatting=True):
        '''with formatting is required for merged cells and border detection'''
        self.wbk = xlrd.open_workbook(filename=filename,
                                      formatting_info=with_formatting)

    def __iter__(self):
        for w in range(self.wbk.nsheets):
            yield xlrdExcelSheet(self.wbk.sheet_by_index(w))
            self.wbk.release_resources()

    def __getitem__(self, name_or_id):
        if isinstance(name_or_id, six.string_types):
            return xlrdExcelSheet(self.wbk.sheet_by_name(name_or_id))
        else:
            return xlrdExcelSheet(self.wbk.sheet_by_index(name_or_id))

load_workbook = xlrdExcelWorkbook
