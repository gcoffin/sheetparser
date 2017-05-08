Results
=======

When a pattern is matched, it fills a `ResultContext`. The
`ResultContext` has to be instantiated by the client and passed to the
match method. Here are the provided classes that derive from
`ResultContext`:

.. autofunction:: sheetparser.results.PythonObjectContext

.. autofunction:: sheetparser.results.ListContext

.. autofunction:: sheetparser.results.DebugContext

