# coding: utf-8
# python 2 and 3 compatibility
from __future__ import print_function
from __future__ import unicode_literals
import re
import abc
from abc import abstractmethod

import six

from .utils import (DoesntMatchException, ConfigurationError, EMPTY_CELL,
                    instantiate_if_class, instantiate_if_class_lst)
from .results import DEFAULT_TRANFORMS
from .documents import (CellRange, WorkbookDocument, SheetDocument,
                        RbRowIterator, RbColIterator, BORDERS_VERTICAL,
                        BORDERS_HORIZONTAL)


def log_matchblock(method):
    def __method(pattern, block, context):
        context.debug(pattern,
                      (not block.empty) and [i.value for i in block.peek],
                      block.idx)
        return method(pattern, block, context)
    return __method


def first_param(fun):
    def __fun(*args):
        return fun(args[0])
    return __fun


def default(default_type, **kwargs):
    (arg_name, default_value), = six.iteritems(kwargs)

    def __decorated(fun):
        def __fun(self, *args, **kwargs):
            if arg_name not in kwargs:
                if len(args) == 0:
                    value = default_value
                elif default_type(args[0]):
                    value = args[0]
                    args = args[1:]
                else:
                    value = default_value
                args = (value,) + args
            else:
                args = (kwargs.pop(arg_name),)+args
            return fun(self, *args, **kwargs)
        return __fun

    return __decorated


def str_or_none(s):
    return s is None or isinstance(s, six.string_types)


@six.add_metaclass(abc.ABCMeta)
class Pattern(object):
    def __repr__(self):
        return "<%s>" % (self.__class__.__name__)

    def assert_type(self, doc):
        pass


@six.add_metaclass(abc.ABCMeta)
class NamedPattern(Pattern):
    def __init__(self, name):
        if not (name is None or isinstance(name, six.string_types)):
            raise ValueError(
                "%s expected name to be a string, got %s" % (self.__class__, name))
        self.name = name

    def __repr__(self):
        return "<%s %s>" % (self.__class__.__name__, self.name)


@six.add_metaclass(abc.ABCMeta)
class BlockPattern(Pattern):
    def __or__(self, p):
        return OrPattern(self, p)

    def __add__(self, p):
        return Sequence('', self, p)

    @abstractmethod
    def match_block(self, iterdoc, context):
        raise NotImplementedError


class OrPattern(BlockPattern):
    """Matches the first pattern and if it fails tries the seconds"""
    def __init__(self, op1, op2):
        self.op1 = instantiate_if_class(op1, BlockPattern)
        self.op2 = instantiate_if_class(op2, BlockPattern)

    def __repr__(self):
        return "<%s | %s>" % (self.op1, self.op2)

    __str__ = __repr__

    @log_matchblock
    def match_block(self, iterdoc, context):
        with iterdoc.rollback_if_fail(reraise=False):
            self.op1.match_block(iterdoc, context)
            return
        self.op2.match_block(iterdoc, context)


class Sequence(NamedPattern, BlockPattern):
    """matches the sub patterns in sequence. All or nothing"""
    @default(str_or_none, name='sequence')
    def __init__(self, name, *patterns):
        self._patterns = list(patterns)
        super(Sequence, self).__init__(name)

    def get_patterns(self):
        return instantiate_if_class_lst(self._patterns, Pattern)

    def emit_meta(self, doc, context):
        pass

    @log_matchblock
    def match_block(self, iterdoc, context):
        with iterdoc.rollback_if_fail(reraise=True):
            with context.push_named(self.name, 'dict'):
                for pattern in self.get_patterns():
                    pattern.match_block(iterdoc, context)

    def __add__(self, pattern):
        self._patterns.append(pattern)
        return self


class Many(NamedPattern, BlockPattern):
    """Matches the subpattern several times. The number
    of times is limited by the parameters max and min"""
    @default(str_or_none, name='many')
    def __init__(self, name, pattern, min=0, max=None):
        self.min = min
        self.max = max
        self.pattern = instantiate_if_class(pattern, Pattern)
        super(Many, self).__init__(name=name)

    def get_patterns(self):
        i = 0
        while True:
            yield "%s%d" % (self.name or '', i), self.pattern
            i += 1

    @log_matchblock
    def match_block(self, iterdoc, context):
        count = 0
        iterpat = self.get_patterns()
        with iterdoc.rollback_if_fail(reraise=True):
            with context.push_named(self.name, 'list'):
                while True:
                    try:
                        with iterdoc.rollback_if_fail():
                            name, pattern = next(iterpat)
                            pattern.match_block(iterdoc, context)
                            count += 1
                            if count == self.max:
                                return
                    except DoesntMatchException:
                        if (self.min > count) or (self.max and self.max < count):
                            # context catches that
                            raise DoesntMatchException(
                                'Bad count (%d) for %s (expected between %s and %s' %
                                (count, self.name, self.min, self.max))
                        break


class Maybe(Many):
    """Matches the subpattern or nothing. Equivalent to ? in
    regexes"""
    @default(str_or_none, name='maybe')
    def __init__(self, name, pattern):
        super(Maybe, self).__init__(name, pattern, min=0, max=1)


class Workbook(Pattern):
    @default(str_or_none, name='workbook')
    def __init__(self, names_dct=None, re_dct=None, *args, **options):
        self.include_hidden = options.get('include_hidden', False)
        self.names_dct = names_dct or {}
        re_dct = re_dct or {}
        if isinstance(re_dct, dict):
            re_dct = six.iteritems(re_dct)
        self.re_list = [(re.compile(r), pattern) for (r, pattern) in re_dct]
        self.seq_patterns = args or None

    def assert_type(self, doc):
        if not isinstance(doc, WorkbookDocument):
            raise ConfigurationError("Expected Workbook, got %s" % doc)

    def match_range_s(self, sheet, pattern_s, context):
        if isinstance(pattern_s, Pattern):
            pattern_s = [pattern_s]
        for pattern in pattern_s:
            pattern.match_range(sheet, context)

    def match_workbook(self, workbook, context):
        """the context will contain the matching sheet in the same order as in the workbook"""
        self.assert_type(workbook)
        with context.push_named('workbook', 'list'):
            if self.seq_patterns:
                patterns_seq = iter(self.seq_patterns)
            names_dct = self.names_dct.copy()
            for s in workbook:
                if s.is_hidden() and not self.include_hidden:
                    continue
                if s.name in names_dct:
                    self.match_range_s(s, names_dct.pop(s.name), context)
                else:
                    for regex, pattern in self.re_list:
                        if regex.match(s.name):
                            self.match_range_s(s, pattern, context)
                            break
                    else:
                        if self.seq_patterns:
                            six.next(patterns_seq).match_range(s, context)
            if self.seq_patterns:
                try:
                    six.next(patterns_seq)
                except StopIteration:
                    pass
                else:
                    raise DoesntMatchException('Some sheets where not visited')


@six.add_metaclass(abc.ABCMeta)
class RangePattern(NamedPattern):
    """Super class for all patterns that match a range"""
    def __init__(self, name, *patterns):
        self._patterns = list(patterns)
        super(RangePattern, self).__init__(name)

    @abstractmethod
    def assert_type(self, doc):
        raise NotImplementedError()

    def get_patterns(self):
        return instantiate_if_class_lst(self._patterns, Pattern)

    def match_range(self, rge, context):
        self.assert_type(rge)
        it = self.iter_range(rge)
        with context.push_named(self.name, 'dict'):
            self.emit_meta(rge, context)
            for pattern in self.get_patterns():
                pattern.match_block(it, context)


@six.add_metaclass(abc.ABCMeta)
class WithLayoutPattern(RangePattern):
    def __init__(self, name, layout, *patterns):
        if isinstance(layout, six.class_types):
            layout = layout()
        if not isinstance(layout, Layout):
            raise ConfigurationError("Expected layout, got %s" % (layout))
        self.layout = layout
        super(WithLayoutPattern, self).__init__(name, *patterns)

    def iter_range(self, rge):
        return self.layout.iter_doc(rge)


class Range(WithLayoutPattern):
    """A range of cells delimited by top, left, bottom,
    right. RangePatterns are to be used directly under Workbook
    """
    def __init__(self, name, layout, *patterns, **kwargs):
        self.top, self.left, self.bottom, self.right = [
            kwargs.pop(n, None) for n in ('top', 'left', 'bottom', 'right')]
        super(Range, self).__init__(name, layout, *patterns)

    def iter_range(self, rge):
        return self.layout.iter_doc(
            CellRange(rge, self.top, self.left, self.bottom, self.right))

    def assert_type(self, doc):
        assert isinstance(doc, CellRange)

    def emit_meta(self, sheet, context):
        context.emit('__meta', {
                'range': (self.top, self.left, self.bottom, self.right)
                })


class Sheet(WithLayoutPattern):
    def assert_type(self, doc):
        if not isinstance(doc, SheetDocument):
            raise ValueError('Expected SheetDocument, got %s' %
                             doc.__class__.__name__)

    def emit_meta(self, sheet, context):
        context.emit('__meta', {'name': sheet.name})


def empty_line(cells):
    return all(cell.is_empty() for cell in cells)


def no_vertical(cells, line_count):  # could do better than that
    """check that there is no vertical line in the cells"""
    return all(not cell.has_borders(BORDERS_VERTICAL) for cell in cells)


def no_horizontal(cells, line_count):
    return all(not cell.has_borders(BORDERS_HORIZONTAL) for cell in cells)


class Table(NamedPattern, BlockPattern):
    """A range of cells read from a line iterator. The table
    transforms are read in sequence at 2 times: when new lines are
    appended and when the table is complete.
    """
    @default(str_or_none, name='table')
    def __init__(self, name, table_args=DEFAULT_TRANFORMS, stop=None):
        self.stop = stop or first_param(empty_line)
        assert callable(self.stop), "stop is not callable: %s" % stop
        self.table_args = table_args
        super(Table, self).__init__(name)

    @log_matchblock
    def match_block(self, iterdoc, context):
        with context.push_named(self.name, 'table'):
            table = context.current
            table.set_args(self.table_args)
            for line_count, g in enumerate(iterdoc):
                table.append_table(g)
                if iterdoc.empty or self.stop(iterdoc.peek, line_count):
                    break
            table.wrap()


class FlexibleRange(WithLayoutPattern):
    """Finds a range by itering through the block until the stop test
    returns true. That range is then used as a new block with the
    given layout abd patterns

    :name: pattern name
    :layout: layout used to iter the result range
    :patterns: patterns to be used with the new layout
    :stop: stop test, by default empty line
    :min: minimum length of the range
    :max: maximum length of the range (None for unbound)
    """
    def __init__(self, name, layout, *patterns, **kwargs):
        self.stop = kwargs.pop('stop', None) or first_param(empty_line)
        self.min = kwargs.pop('min', 1)
        self.max = kwargs.pop('max', None)
        super(FlexibleRange, self).__init__(name, layout, *patterns)

    def assert_type(self, doc):
        assert isinstance(doc, (CellRange, RbRowIterator, RbColIterator))

    def __repr__(self):
        return "<%s %s>" % (self.__class__.__name__, self.name)

    @log_matchblock
    def match_block(self, iterdoc, context):
        if iterdoc.empty or empty_line(iterdoc.peek):
            raise DoesntMatchException()
        s = iterdoc.peek
        top, left, bottom, right = s.top, s.left, s.bottom, s.right
        linecount = 0
        for linecount, g in enumerate(iterdoc):
            top = min(top, g.top)
            left = min(left, g.left)
            bottom = max(bottom, g.bottom)
            right = max(right, g.right)
            if iterdoc.empty or self.stop(iterdoc.peek, linecount):
                break
        if self.min > linecount:
            raise DoesntMatchException(
                'Flexible range %s has %d lines, min is %d' % (self.name, linecount, self.min))
        if self.max is not None and linecount > self.max:
            raise DoesntMatchException(
                'Flexible range %s has %d lines, max is %d' % (self.name, linecount, self.max))
        rge = CellRange(s.rge, top, left, bottom, right)
        super(FlexibleRange, self).match_range(rge, context)

    def emit_meta(self, sheet, context):
        context.emit('__meta', {'flexible': self.name})


class Line(NamedPattern, BlockPattern):
    """Matches a line: there must be one more row/column in the block
    and it must be non empty.
    :line_args: list of transforms to the result (strip, raise if empty...)
    """
    @default(str_or_none, name='line')
    def __init__(self, name, line_args=None):
        super(Line, self).__init__(name)
        self.line_args = line_args or []

    @log_matchblock
    def match_block(self, iterdoc, context):
        if iterdoc.empty:
            raise DoesntMatchException("Line %s does not match (end of block)" % self.name)
        with context.push_named(self.name,'line'):
            line = context.current
            line.set_args(self.line_args)
            line.set_value(iterdoc.peek)
        six.next(iterdoc)


class Empty(BlockPattern):
    """Matches an empty line. Doesn't match if there is no more lines
    in the block
    """

    @log_matchblock
    def match_block(self, iterdoc, context):
        if iterdoc.empty:
            raise DoesntMatchException('%s expects a line' %
                                       (self, ))
        if not empty_line(iterdoc.peek):
            raise DoesntMatchException('%s not matched by %s' %
                                       (self, list(iterdoc.peek)))
        else:
            six.next(iterdoc)


# Layouts


@six.add_metaclass(abc.ABCMeta)
class Layout(object):
    @abstractmethod
    def iter_doc(self, doc):
        pass


class Rows(Layout):
    def iter_doc(self, doc):
        assert isinstance(doc, CellRange), "Expected CellRange, got %s" % doc
        return RbRowIterator(doc)


class Columns(Layout):
    def iter_doc(self, doc):
        assert isinstance(doc, CellRange), "Expected CellRange, got %s" % doc
        return RbColIterator(doc)
