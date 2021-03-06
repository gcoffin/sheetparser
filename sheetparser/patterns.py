# coding: utf-8
import abc
import re
from abc import abstractmethod
import warnings

from .documents import (CellRange, WorkbookDocument, SheetDocument,
                        RbRowIterator, RbColIterator, RbVisibleRowIterator,
                        BORDERS_VERTICAL,
                        BORDERS_HORIZONTAL)
from .results import DEFAULT_TRANSFORMS
from .utils import (DoesntMatchException, ConfigurationError,
                    instantiate_if_class, instantiate_if_class_lst)


def log_match_iterator(method):
    def __method(pattern, line_iterator, context):
        context.debug(pattern,
                      ('<no line>' if line_iterator.is_complete
                       else [i.value for i in line_iterator.peek]),
                      'Idx:', line_iterator.idx)
        return method(pattern, line_iterator, context)

    return __method


def first_param(fun):
    """drops all arguments except the first one then calls the
    decorated function"""

    def __fun(*args):
        return fun(args[0])

    return __fun


def default(default_type, **kwargs):
    """A decorator that assigns a default value to the
    first argument of a function if it doesn't match the
    default_type function"""
    (arg_name, default_value), = kwargs.items()

    def __decorated(fun):
        def __fun(self, *args, **kwargs):
            if arg_name not in kwargs:
                if len(args) == 0:
                    value = default_value
                elif isinstance(args[0], default_type):
                    warnings.warn(
                        f"the default decorated is deprecated, "
                        f"use {arg_name}={repr(default_value)} instead"
                    )
                    value = args[0]
                    args = args[1:]
                else:
                    value = default_value
                kwargs[arg_name] = value
            return fun(self, *args, **kwargs)

        return __fun

    return __decorated


str_or_none = str


class Pattern(abc.ABC):
    def __repr__(self):
        return "<%s>" % (self.__class__.__name__)

    def assert_type(self, doc):
        pass


class NamedPattern(Pattern, metaclass=abc.ABCMeta):
    def __init__(self, name):
        if not (name is None or isinstance(name, str)):
            raise ValueError(
                "%s expected name"
                " to be a string, got %s" % (self.__class__, name))
        self.name = name

    def __repr__(self):
        return "<%s %s>" % (self.__class__.__name__, self.name)


class LineIteratorPattern(Pattern, metaclass=abc.ABCMeta):
    def __or__(self, p):
        return OrPattern(self, p)

    def __add__(self, p):
        return Sequence('add_', self, p)

    @abstractmethod
    def match_line_iterator(self, line_iterator, context):
        raise NotImplementedError


class AbstractRangePattern(Pattern, metaclass=abc.ABCMeta):
    def __or__(self, p):
        return RangeOr(self, p)

    def __add__(self, p):
        return RangeAnd(self, p)

    @abstractmethod
    def match_range(self, range, context):
        raise NotImplementedError


class RangeOr(NamedPattern, AbstractRangePattern):
    @default(str_or_none, name='or')
    def __init__(self, pattern1, pattern2, name='or'):
        self.pattern1 = instantiate_if_class(pattern1, AbstractRangePattern)
        self.pattern2 = instantiate_if_class(pattern2, AbstractRangePattern)
        super(RangeOr, self).__init__(name)

    def __repr__(self):
        return "<%s | %s>" % (self.pattern1, self.pattern2)

    __str__ = __repr__

    def match_range(self, range, context):
        try:
            with context.push_named(self.name, 'dict'):
                self.pattern1.match_range(range, context)
                return
        except DoesntMatchException:
            self.pattern2.match_range(range, context)


class RangeAnd(NamedPattern, AbstractRangePattern):
    @default(str_or_none, name='and_')
    def __init__(self, *patterns, name='and_'):
        self._patterns = list(patterns)
        super(RangeAnd, self).__init__(name)

    def get_patterns(self):
        return instantiate_if_class_lst(self._patterns, AbstractRangePattern)

    def emit_meta(self, doc, context):
        pass

    def match_range(self, range, context):
        with context.push_named(self.name, 'dict'):
            for pattern in self.get_patterns():
                pattern.match_range(range, context)

    def __add__(self, pattern):
        self._patterns.append(pattern)
        return self


class OrPattern(NamedPattern, LineIteratorPattern):
    """matches the first pattern and if it fails tries the seconds.

    :param Pattern pattern1: first pattern to try
    :param Pattern pattern2: fall back patter
    """

    @default(str_or_none, name='or_')
    def __init__(self, *args, name='or_'):
        super().__init__(name)
        self.patterns = [instantiate_if_class(p, LineIteratorPattern)
                         for p in args]

    def __repr__(self):
        return "<%s>" % (' | '.join(str(i) for i in self.patterns))

    __str__ = __repr__

    @log_match_iterator
    def match_line_iterator(self, line_iterator, context):
        for pattern in self.patterns:
            with line_iterator.rollback_if_fail(reraise=False):
                # ignore what was pushed in the context in case of failure
                with context.push_named(self.name, 'dict'):
                    pattern.match_line_iterator(line_iterator, context)
                    return
        raise DoesntMatchException


class Sequence(NamedPattern, LineIteratorPattern):
    """matches the sub patterns in sequence. Will match all or nothing.
    Name is an optional parameter. If omitted, the name will be 'sequence'.
    """

    @default(str_or_none, name='sequence')
    def __init__(self, *patterns, name='sequence'):
        """this is docstring

        :param str name: the name
        """
        self._patterns = list(patterns)
        super(Sequence, self).__init__(name)

    def get_patterns(self):
        return instantiate_if_class_lst(self._patterns, Pattern)

    def emit_meta(self, doc, context):
        pass

    @log_match_iterator
    def match_line_iterator(self, line_iterator, context):
        with line_iterator.rollback_if_fail(reraise=True):
            with context.push_named(self.name, 'dict'):
                for pattern in self.get_patterns():
                    pattern.match_line_iterator(line_iterator, context)

    def __add__(self, pattern):
        self._patterns.append(pattern)
        return self


class Many(NamedPattern, LineIteratorPattern):
    """Matches the subpattern several times. The number of times is
    limited by the parameters max and min. Name defaults to 'many'"""

    @default(str_or_none, name='many')
    def __init__(self, pattern, min=0, max=None, name='many'):
        if not ((min is None or isinstance(min, int)) and
                (max is None or isinstance(max, int))):
            raise ConfigurationError('In Many, min and max need to be numbers')
        self.min = min
        self.max = max
        self.pattern = instantiate_if_class(pattern, Pattern)
        super(Many, self).__init__(name=name)

    def get_patterns(self):
        i = 0
        while True:
            yield "%s%d" % (self.name or '', i), self.pattern
            i += 1

    @log_match_iterator
    def match_line_iterator(self, line_iterator, context):
        count = 0
        iterpat = self.get_patterns()
        with line_iterator.rollback_if_fail(reraise=True):
            with context.push_named(self.name, 'list'):
                while True:
                    try:
                        with line_iterator.rollback_if_fail():
                            name, pattern = next(iterpat)
                            pattern.match_line_iterator(line_iterator, context)
                            count += 1
                            if count == self.max:
                                return
                    except DoesntMatchException:
                        if (self.min > count) or (self.max and
                                                  self.max < count):
                            # context catches that
                            raise DoesntMatchException(
                                'Bad count (%d) for %s'
                                ' (expected between %s and %s' %
                                (count, self.name, self.min, self.max))
                        break


class Maybe(Many):
    """Matches the subpattern or nothing. Equivalent to ? in
    regexes"""

    @default(str_or_none, name='maybe')
    def __init__(self, pattern, name='maybe'):
        super(Maybe, self).__init__(name, pattern, min=0, max=1)


class Workbook(NamedPattern):
    """A top level pattern to match a workbook. Call match_workbook on
    an opened workbook document (as provided by a backend)

    :param dict_or_list patterns: a list of patterns or a dictionary
        that associates a sheet name or regular expession to the sheet
        pattern

    :param dict regex: dictionary that associates a regular expression
        to a pattern. If a sheet matches the regex, then it will try
        to match the pattern.
    """

    @default(str_or_none, name='workbook')
    def __init__(self, patterns=None, **options):
        self.name=options.pop('name')
        super().__init__(self.name)
        self.include_hidden = options.get('include_hidden', False)
        self.seq_patterns = ()
        self.names_dct = {}
        # for backward compatibility
        re_iter = options.pop('regex', None) or options.pop('re_dct', ())
        if isinstance(re_iter, dict):
            re_iter = re_iter.items()
        self.re_list = [(re.compile(r), pattern) for (r, pattern) in re_iter]
        if isinstance(patterns, dict):
            self.names_dct = patterns
        else:
            self.seq_patterns = patterns

    def assert_type(self, doc):
        if not isinstance(doc, WorkbookDocument):
            raise ConfigurationError("Expected Workbook, got %s" % doc)

    def _match_range_s(self, sheet, pattern_s, context):
        context.debug('sheet', sheet.name)
        if isinstance(pattern_s, Pattern):
            pattern_s = [pattern_s]
        for pattern in pattern_s:
            pattern.match_range(sheet, context)

    def match_workbook(self, workbook, context):
        """Iterates through the sheets in the workbook. If `names_dct`
        contains the sheet name, it will try and match the associated
        pattern. If not, the method will try in `re_dct` if any of the
        regular expressions matches the names. Finally, if any other
        pattern is provided, they will be tried in sequence.

        The context will contain the matching sheet in the same order
        as in the workbook, """
        self.assert_type(workbook)
        with context.push_named('workbook', 'list'):
            if self.seq_patterns:
                patterns_seq = iter(self.seq_patterns)
            names_dct = self.names_dct.copy()
            for s in workbook:
                if s.is_hidden() and not self.include_hidden:
                    continue
                context.debug('workbook', repr(s.name), names_dct)
                if s.name in names_dct:
                    self._match_range_s(s, names_dct.pop(s.name), context)
                else:
                    for regex, pattern in self.re_list:
                        if regex.match(s.name):
                            self._match_range_s(s, pattern, context)
                            break
                    else:
                        if self.seq_patterns:
                            next(patterns_seq).match_range(s, context)
            if self.seq_patterns:
                try:
                    next(patterns_seq)
                except StopIteration:
                    pass
                else:
                    raise DoesntMatchException('Some sheets where not visited')


class RangePattern(NamedPattern, AbstractRangePattern, metaclass=abc.ABCMeta):
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
                pattern.match_line_iterator(it, context)


class WithLayoutPattern(RangePattern, metaclass=abc.ABCMeta):
    def __init__(self, name, layout, *patterns):
        if isinstance(layout, type):
            layout = layout()
        if not isinstance(layout, Layout):
            raise ConfigurationError("Expected layout, got %s" % (layout))
        self.layout = layout
        super(WithLayoutPattern, self).__init__(name, *patterns)

    def iter_range(self, rge):
        return self.layout.iter_doc(rge)


class Range(WithLayoutPattern):
    """A range of cells delimited by top, left, bottom,
    right. RangePatterns are to be used directly under Workbook.
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
            'range': (self.top, self.left, self.bottom, self.right),
            'name': sheet.name,
        })


class Sheet(WithLayoutPattern):
    def assert_type(self, doc):
        if not isinstance(doc, SheetDocument):
            raise ValueError('Expected SheetDocument, got %s' %
                             doc.__class__.__name__)

    def emit_meta(self, sheet, context):
        context.emit('__meta', {'name': sheet.name})


def empty_line(cells, line_count=0):
    """returns true if all cells are empty"""
    return all(cell.is_empty for cell in cells)


def no_vertical(cells, line_count=0):  # could do better than that
    """check that there is no vertical line in the cells"""
    return all(not cell.has_borders(BORDERS_VERTICAL) for cell in cells)


def no_horizontal(cells, line_count=0):
    """return True is no cell has horizontal border"""
    return all(not cell.has_borders(BORDERS_HORIZONTAL) for cell in cells)


def make_stop_function(stop):
    if stop is None: 
        return None
    if isinstance(stop, str):
        stop = re.compile(stop)
    if isinstance(stop, re.Pattern):
        regex = stop
        stop = lambda line, _, regex=regex: regex.match(''.join(line))
    return stop


class Table(NamedPattern, LineIteratorPattern):
    """A range of cells read from a line iterator. The table
    transforms are read in sequence at 2 times: when new lines are
    appended and when the table is complete.

    :param str name: optional name of the table, "table" by default.
    :param list table_args: the arguments that are sent to the
        ResultContext that will store the result. For ResultTable, the
        default, that will be the list of transforms.
    :param function stop: that function is called on the following
        line. The table end is reached when that function returns
        True. It takes 2 parameters: the number of lines read so far
        and the line itself. By default, will stop on empty lines
    """

    @default(str_or_none, name='table')
    def __init__(self, table_args=DEFAULT_TRANSFORMS, stop=None, name='table'):
        self.stop = make_stop_function(stop) or first_param(empty_line)
        assert callable(self.stop), "stop is not callable: %s" % stop
        self.table_args = table_args
        super(Table, self).__init__(name)

    @log_match_iterator
    def match_line_iterator(self, line_iterator, context):
        with context.push_named(self.name, 'table'):
            table = context.current
            table.set_args(self.table_args)
            for line_count, g in enumerate(line_iterator):
                table.append_table(g)
                if (line_iterator.is_complete or self.stop(
                        line_iterator.peek, line_count)):
                    break
            table.wrap()


class FlexibleRange(WithLayoutPattern):
    """Finds a range by itering through the lines until the stop test
    returns true. That range is then used as a new range with the
    given layout and patterns.

    :param str name: pattern name
    :param Layout layout: layout used to iter the result range
    :param Pattern patterns: patterns to be used with the new layout
    :param function(line_count, line) stop: stop test, by default empty line
    :param int min: minimum length of the range
    :param int max: maximum length of the range (None for unbound)
    """

    @default(str_or_none, name='flexible')
    def __init__(self, layout, *patterns, **kwargs):
        self.stop = make_stop_function(kwargs.pop('stop', None)) or first_param(empty_line)
        self.min = kwargs.pop('min', 1)
        self.max = kwargs.pop('max', None)
        name = kwargs.pop('name')
        super(FlexibleRange, self).__init__(name, layout, *patterns)

    def assert_type(self, doc):
        assert isinstance(doc, (CellRange, RbRowIterator, RbColIterator))

    def __repr__(self):
        return "<%s %s>" % (self.__class__.__name__, self.name)

    @log_match_iterator
    def match_line_iterator(self, line_iterator, context):
        if line_iterator.is_complete:
            raise DoesntMatchException(
                "FlexibleRange didn't match (no more lines %d)" %
                (line_iterator.idx))
        s = line_iterator.peek
        top, left, bottom, right = s.top, s.left, s.bottom, s.right
        linecount = 0
        for linecount, g in enumerate(line_iterator):
            top = min(top, g.top)
            left = min(left, g.left)
            bottom = max(bottom, g.bottom)
            right = max(right, g.right)
            if (line_iterator.is_complete or self.stop(
                    line_iterator.peek, linecount)):
                break
        if self.min > linecount:
            raise DoesntMatchException(
                'Flexible range %s has %d lines, min is %d' %
                (self.name, linecount, self.min))
        if self.max is not None and linecount > self.max:
            raise DoesntMatchException(
                'Flexible range %s has %d lines, max is %d' %
                (self.name, linecount, self.max))
        rge = CellRange(s.rge, top, left, bottom, right)
        context.debug('FlexibleRange', rge)
        super(FlexibleRange, self).match_range(rge, context)

    def emit_meta(self, sheet, context):
        context.emit('__meta', {'flexible': self.name})


class Line(NamedPattern, LineIteratorPattern):
    """Matches a line: there must be one more row/column in the
    line_iterator and it must be non empty.

    :param list line_args: list of transforms to the result (strip,
    raise if empty...)
    """

    @default(str_or_none, name='line')
    def __init__(self, line_args=None, name=''):
        super(Line, self).__init__(name)
        self.line_args = line_args or []

    @log_match_iterator
    def match_line_iterator(self, line_iterator, context):
        if line_iterator.is_complete:
            raise DoesntMatchException(
                "Line %s does not match (end of line_iterator)" % self.name)
        with context.push_named(self.name, 'line'):
            line = context.current
            line.set_args(self.line_args)
            line.set_value(line_iterator.peek)
        next(line_iterator)


class Empty(LineIteratorPattern):
    """Matches an empty line. Doesn't match if there is no more lines
    in the line_iterator
    """

    @log_match_iterator
    def match_line_iterator(self, line_iterator, context):
        if line_iterator.is_complete:
            raise DoesntMatchException('%s expects a line' %
                                       (self,))
        if not empty_line(line_iterator.peek):
            raise DoesntMatchException('%s not matched by %s' %
                                       (self, list(line_iterator.peek)))
        else:
            next(line_iterator)


# Layouts


class Layout(abc.ABC):
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


class VisibleRows(Layout):
    def iter_doc(self, doc):
        assert isinstance(doc, CellRange), "Expected CellRange, got %s" % doc
        return RbVisibleRowIterator(doc)
