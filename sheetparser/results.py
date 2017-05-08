# coding: utf-8
# python 2 and 3 compatibility
from __future__ import print_function
from __future__ import unicode_literals
import datetime
import abc
import re

import six

from .utils import DoesntMatchException, EMPTY_CELL, ConfigurationError, instantiate_if_class_lst


class ResultContext(object):
    '''An object that is passed through match methods to emit the
    result. Implement emit is a concrete subclass'''
    def __init__(self):
        self.root = None
        self.stack = []

    def push(self, level):
        if not self.stack:
            self.root = level
        self.stack.append(level)
        return self

    @property
    def current(self):
        return self.stack[-1]

    def pop(self):
        self.stack.pop()

    def emit(self, name, o):
        raise NotImplementedError()

    def commit(self, o1, o2):
        raise NotImplementedError()

    def debug(self,*args):
        pass

    def __enter__(self):
        return self

    def __exit__(self, etype, evalue, tb):
        if etype is None and len(self.stack) >= 2:
            self.commit(self.stack[-2], self.stack[-1])
            self.stack.pop()
        else:
            self.pop()
        return False


@six.add_metaclass(abc.ABCMeta)
class AbstractVisitor(object):
    def __call__(self, o):
        type_name = type(o).__name__
        if type_name in self.dispatch:
            return getattr(self, self.dispatch[type_name])(o)
        elif hasattr(o, 'visit'):
            return o.visit(self)
        else:
            return o


def table_str(table):
    return '\n'.join(', '.join("%s" % i for i in l) for l in table)


class QuickPrint(AbstractVisitor):
    dispatch = {
        'ResultTable': 'visit_table_with_header'
        }

    def visit_table(self, o):
        return str(o)

    def visit_table_with_header(self, o):
        return {'_header': table_str(o.top_headers),
                '_column': table_str(o.left_headers),
                '_top_left': table_str(o.top_left),
                '_data': table_str(o.data)
                }


# Classes for PythonObjectContext
class ResultObject(object):
    def __init__(self, name, *args, **kwargs):
        self.name = name
        super(ResultObject, self).__init__(*args, **kwargs)

    def set_args(self, *args):
        pass


class ResultDict(ResultObject, dict):
    def visit(self, visitor):
        return {k: visitor(v) for (k, v) in six.iteritems(self)}

    def add(self, name, value):
        suffix = None
        while True:
            n = (name if suffix is None
                 else (name + '_%d' % suffix))
            if n not in self:
                break
            suffix = (suffix or 0) + 1
        self[n] = value

    def __getattr__(self, name):
        return self[name]

    def __repr__(self):
        return "Dict %s (%s)" % (
            self.name,dict.__repr__(self))

class ResultList(ResultObject, list):
    def visit(self, visitor):
        method = getattr(visitor, 'visit_list', None)
        if method is None:
            return [visitor(x) for x in self]
        else:
            return method(self)

    def add(self, name, value):
        self.append(value)

    def __repr__(self):
        return ("List %s (%s)" %
                (self.name,
                 ', '.join(str(i) for i in self)))


def _rindex(lst, x):
    """reverse index (index of first element from the end)"""
    return len(lst) - 1 - lst[::-1].index(x)


class StripLine(object):
    '''A transformer used by Lines to remove trailing and ending empty
    cells
    '''
    def __init__(self, left=True, right=True):
        self.left = left
        self.right = right

    def __call__(self, line):
        empties = [cell.is_empty() for cell in line]
        if all(empties):
            return []
        if self.right:
            line = line[:_rindex(empties, 0)+1]
        if self.left:
            line = line[empties.index(0):]
        return line


def non_empty(line):
    '''A transformer that matches only non empty lines. Other will
    raise a DoesntMatchException'''
    if not line:
        raise DoesntMatchException('Empty line')
    return line


class Match(object):
    '''A transformer that matches lines that contain the given
    regex. Use combine to decide if all or any item should match

    :param regex regex: a regular expression
    :param list position: a list of positions or a slice
    :param function combine: function that decides if the whole line
        matches
    '''
    def __init__(self, regex, position=None, combine=None):
        self.regex = re.compile(regex)
        self.combine = None or any
        if isinstance(position, six.integer_types):
            self.position = [position]
        elif position is None:
            self.position = slice(None,None)
        else:
            self.position = position

    def __call__(self, line):
        if isinstance(self.position, slice):
            if not self.combine(self.regex.match(p) for p in line[self.position]):
                raise DoesntMatchException
        elif not self.combine(self.regex.match(line[p]) for p in self.position):
            raise DoesntMatchException
        return line


def get_value(line):
    '''A transformer that converts a list of cells to a list of values'''
    return [c.value if not c.is_merged else EMPTY_CELL for c in line]


class ResultLine(ResultObject, list):
    def set_args(self, transforms=None):
        self._transforms = transforms or [StripLine(), non_empty, get_value]

    def visit(self, visitor):
        visitor.visit_line(self)

    def set_value(self,line):
        line = list(line)
        for t in self._transforms:
            line = t(line)
        self[:] = line

class ResultTable(ResultObject):
    '''An object to store the content of a matched Table.
This is a'''
    def __init__(self, name, transforms=None , iffail='no match'):
        self.name = name
        self.data = []
        self.count = 0
        self.set_args(transforms, iffail)

    def set_args(self, transforms=None , iffail='no match'):
        self.transforms = instantiate_if_class_lst(transforms or [], TableTransform)
        self.iffail = {'no match': DoesntMatchException,
                       'fail': None}[iffail]

    def append_table(self, line):
        for transform in self.transforms:
            line = transform.process_line(self, line)
            if line is None:
                break
        self.count += 1

    def wrap(self):
        for transform in self.transforms:            
            try:
                transform.wrap(self)
            except Exception as e:
                if self.iffail is not None:
                    raise
                    six.raise_from(DoesntMatchException, e)
                else:
                    raise

    def __repr__(self):
        return "Table %s (%s)" % (self.name, self.data)


class PythonObjectContext(ResultContext):
    """Store the results are a hierarchy of objects that mimics the
    initial hierarchy of patterns"""
    types = {'list': ResultList,
             'dict': ResultDict,
             'line': ResultLine, 
             'table': ResultTable}

    def __init__(self):
        super(PythonObjectContext, self).__init__()

    def push_named(self, name, type_):
        if type_ is None:
            o = self.current  # !!! won't work with rollback
        else:
            o = self.types[type_](name=name)
        return super(PythonObjectContext, self).push(o)

    def emit(self, name, o):
        self.current.add(name, o)

    def commit(self, o1, o2):
        o1.add(o2.name, o2)

    def __getattr__(self, name):
        return getattr(self.root, name)

    def __getitem__(self, name):
        return self.root[name]


class ListContext(PythonObjectContext):
    '''a context that returns a dictionary where the key is the name
    of the pattern'''
    class DefaultResult(dict):
        def __init__(self,name):
            self.name = name
            dict.__init__(self)
        def append(self,arg):
            name, value = arg
            self.setdefault(name,[]).append(value)

    types = { 'list': DefaultResult,
              'dict': DefaultResult,
              'line': ResultLine,
              'table': ResultTable }

    def emit(self, name, o):
        self.current.append((name, o))

    def commit(self, o1, o2):
        if isinstance(o2,ListContext.DefaultResult):
            o1.update(o2)
        else:
            o1.append((o2.name,o2))


class DebugContext(ListContext):
    '''A result context that implements the debug function'''
    def debug(self,*args):
        print(' '*len(self.stack),*args)

    def pop(self):
        print(' '*len(self.stack),'--')
        super(DebugContext,self).pop()

    def commit(self,*args):
        self.debug('++')
        super(DebugContext,self).commit(*args)

class TableTransform(object):
    def wrap(self, table):
        pass

    def process_line(self, table, line):
        return line


class TableNotEmpty(TableTransform):
    def process_line(self, table, line):
        if not any(line): return None
        return line

    def wrap(self, table):
        if not table.data:
            raise DoesntMatchException('TableNotEmpty failed: No data in table')

class GetValue(TableTransform):
    """Transforms a list of cells into a list of strings. All built in
    processors expect GetValue to be included as the first
    transformation."""
    def process_line(self,table,line):
        return [x.value for x in line]

class FillData(TableTransform):
    """Adds the line to the table data"""
    def process_line(self, table, line):
        table.data.append(line)


class HeaderTableTransform(TableTransform):
    """Extract the first lines and first columns
    as the top and left headers

    :param int top_header: number of lines, 1 by default
    :param int left_column: number of columns, 1 by default
    """
    def __init__(self, top_header=1, left_column=1):
        self.top_header = top_header
        self.left_column = left_column

    def process_line(self, table, line):
        if not line:
            return
        if table.count == 0:
            table.top_headers = []
            table.left_headers = []
        if self.left_column:
            col = line[:self.left_column]
            line = line[self.left_column:]
            if table.count == 0:
                table.left_headers = [[] for i in range(self.left_column)]
                table.top_left = [[] for i in range(self.left_column)]
            if table.count >= self.top_header:
                for h, c in zip(table.left_headers, col):
                    h.append(c)
            else:
                for h, c in zip(table.top_left, col):
                    h.append(c)
        if self.top_header and table.count < self.top_header:
            table.top_headers.append(line)
        else:
            return line
        return None


class RepeatExisting(TableTransform):
    '''Replaces empty strings with previous data'''
    def wrap(self, table):
        table.top_headers = [_repeat_existing(i) for i in table.top_headers]


def _find_non_empty_rows(list_of_lists):
    return [i for i, line in enumerate(list_of_lists)
            if any(x != EMPTY_CELL for x in line)]

class RemoveEmptyLines(TableTransform):
    '''Remove empyt lines or empty columns in the table. Note: could
    be really simplified with numpy'''
    def __init__(self, line_type='rows'):
        if line_type not in ['rows', 'columns']:
            raise ConfigurationError(
                "line_type must be 'rows' or 'columns' - got %s"
                % repr(line_type))
        self.line_type = line_type

    def wrap(self, table):
        if self.line_type == 'columns':
            Transpose().wrap(table)
        data_rows = _find_non_empty_rows(table.data)
        table.data = [table.data[i] for i in data_rows] 
        if hasattr(table,'left_headers'):
            tlf = transpose(table.left_headers)
            table.left_headers = transpose(tlf[i] for i in data_rows)
        if self.line_type == 'columns':
            Transpose().wrap(table)


class ToMap(TableTransform):
    """Transforms the data from a list of lists to a map. The keys are
    the combination of terms in the headers (top and left) and the
    values are the table data"""
    def wrap(self, table):
        result = {}
        for lefts, row in zip(zip(*table.left_headers), table.data):
            for tops, cell in zip(zip(*table.top_headers), row):
                key = tuple(lefts)+tuple(tops)
                result[key] = cell
        table.data = result
        

def _join_header(lines, char):
    return [char.join("%s" % s for s in u) for u in zip(*lines)]


class MergeHeader(TableTransform):
    """merges several lines in the header into one"""
    def __init__(self, join_top=(), join_left=(), ch='.'):
        if not all(isinstance(i, int) for i in join_top):
            raise ConfigurationError('ids must be ints, got %s' % join_top)
        if not all(isinstance(i, int) for i in join_left):
            raise ConfigurationError('ids must be ints, got %s' % join_left)
        self.join_char = ch
        self.join_left = join_left
        self.join_top = join_top

    def merge(self, header, ids):
        to_merge = [header[i] for i in ids]
        not_merge = [h for i, h in enumerate(header) if i not in ids]
        to_merge = [_join_header(to_merge, self.join_char)]
        return to_merge + not_merge

    def wrap(self, table):
        if self.join_top:
            table.top_headers = self.merge(table.top_headers, self.join_top)
        if self.join_left:
            table.left_headers = self.merge(table.left_headers, self.join_left)


def transpose(list_of_lists):
    return list(list(r) for r in zip(*list_of_lists))


class Transpose(TableTransform):
    """Transforms lines into columns and columns to lines"""
    def wrap(self, table):
        if hasattr(table,'top_headers') and hasattr(table,'left_headers'):
            table.top_headers, table.left_headers = table.left_headers, table.top_headers
        table.data = transpose(table.data)


class ToDate(TableTransform):
    """Transforms strings into dates in the header. Use merge if the
    date is spread over several lines"""
    def __init__(self, header_id, strftime, is_top=True, join='/'):
        self.header_id = header_id
        self.is_top = is_top
        if isinstance(strftime, six.string_types):
            self.strftime = (lambda s, format=strftime:
                             datetime.datetime.strptime(s, format))
        else:
            self.strftime = strftime
        self.join = join

    def wrap(self, table):
        headers = table.top_headers if self.is_top else table.left_headers
        dates_str = headers.pop(self.header_id)
        result = []
        for d in dates_str:
            try:
                result.append(self.strftime(d))
            except ValueError:
                result.append(d)
        headers.append(result)


def _repeat_existing(line):
    current = None
    result = []
    for i in line:
        if i != EMPTY_CELL:
            current = i
        result.append(current)
    return result


DEFAULT_TRANFORMS = [GetValue, HeaderTableTransform, FillData, TableNotEmpty]
