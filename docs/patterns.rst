Patterns
********

.. currentmodule :: sheetparser.patterns

Some definitions
----------------

    range 
        is an Excel range, delimited with a top, a left, a right and a
        bottom. A sheet is an example of a range.
    
    line
        is a row or a column. This is decided by the chosen **layout**:
        horizontal layouts will yield rows, vertical will yield columns.
    
    pattern
        is an object that matches the given range or line(s). If
        the match fails, the method raises a DoesntMatchException. If it
        succeeds, it fills up the context given as a parameter. 
    
    Note that patterns can be passed as arguments to the upper level
    pattern as object or classes. Classes will be instatianted.

There are 3 types of patterns: 

Workbook 
---------
This pattern will be called to match a workbook:

.. autoclass:: Workbook(names_dct=None, re_dct=None, *args, **options)
    :members:

Ranges
------

The following patterns match either the whole sheet or a
range:

.. autoclass:: Sheet(name, layout, *patterns)

.. autoclass:: Range(name, layout, *patterns, top=None, left=None, bottom=None, right=None)

Layout is Rows or Columns, and will be used to know if the range
should be read horizontally or vertically.

Iterators of lines
------------------

These patterns are called on an iterator of lines, and will be passed
as parameters to Range patterns or other patterns matching iterators
of lines.

These patterns can be combined with the operator +, which returns a
Sequence. a+b is equivalent to Sequence(a,b). Similarly, a|b is
equivalent to OrPattern(a,b).

The name of the pattern is used by the ResultContext to store the
matched element.  The existing patterns that operate on an line
iterator are:

.. autoclass:: Empty(name)

.. autoclass:: Sequence(name='sequence',*patterns)

.. autoclass:: Many(name='many',pattern)

.. autoclass:: Maybe(name=None,pattern)

.. autoclass:: OrPattern(pattern1, pattern2)

.. autoclass:: FlexibleRange(name='flexible', layout, *patterns, stop=None, min=None, max=None)

.. autoclass:: Table(name='table', table_args=DEFAULT_TRANFORMS, stop=None)

.. autoclass:: Line(name='line', line_args=None)

Stop tests
----------

Stop tests are functions that are passed to FlexibleRange and Table to detect the end of a block.

.. autofunction:: empty_line(cells, line_count)

.. autofunction:: no_horizontal

.. autofunction:: no_vertical
