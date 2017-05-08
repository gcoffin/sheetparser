Line and table transformations
********************************

The contents that is matched by the `Line` and `Table` patterns is stored
in the context result. Another level of processing is provided by list
of transformations.

Line transformations
--------------------

They are passed as *line_args* parameters to the Line pattern. It is a
list of function that take a list and return a list. These functions
are called in sequence, the result of one function is passed to the
following one. 

The first function of the list must accept a list of Cell. The
function get_value transforms it to the list of values.

These are the included line transformations:

.. autofunction:: sheetparser.results.non_empty(line)

Parameterized functions (objects with a method __call__):

.. autoclass:: sheetparser.results.StripLine(left=True, right=True)

.. autoclass:: sheetparser.results.Match(regex, position=None, combine=None)


Table transformations
---------------------

Similarly, the lines matched by the `Table`  pattern are passed to a
series of processings. They are subclasses of `TableTransform` which
implement `wrap` or `process_line` (or both). `process_line` is called
when a new line is added, and `wrap` is called at the end when all
lines have been added.

.. autoclass:: sheetparser.results.GetValue

.. autoclass:: sheetparser.results.FillData

.. autoclass:: sheetparser.results.HeaderTableTransform

.. autoclass:: sheetparser.results.RepeatExisting

.. autoclass:: sheetparser.results.RemoveEmptyLines

.. autoclass:: sheetparser.results.ToMap

.. autoclass:: sheetparser.results.MergeHeader

.. autoclass:: sheetparser.results.Transpose
