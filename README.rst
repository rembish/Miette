=================================================
Python Library for reading .doc MS Word documents
=================================================

|PyPI latest| |PyPI Py Versions|

Miette is a "small sweet thing" in French.

In another way, Miette is light-weight low memory usage library for
reading and, maybe converting (in future) Microsoft Word 97-2003 documents 
(.doc).

Thanks to contributor(s):
Jason Brechin <brechinj[at]gmail[daht]com

It'll be last version of Miette 1.x project. Now this project has readable
code, better exception management, but has no future in current state. I'll
try to start child project with better semantic and logic. See ya soon.

With best wishes, yours Alex Rembish.

Usage
-----

.. code-block:: python

   >>> import miette
   >>> bs = miette.DocReader('example.doc').read()
   >>> print(bs.decode('utf-8'))


Related
-------
Python libraries for reading .docx Microsoft Word files: 

- `python-docx2txt <https://github.com/ankushshah89/python-docx2txt>`_
- `python-docx <https://github.com/python-openxml/python-docx>`_


.. |PyPI latest| image:: https://img.shields.io/pypi/v/miette.svg?maxAge=360
   :target: https://pypi.python.org/pypi/miette
.. |PyPI Py Versions| image:: https://img.shields.io/pypi/pyversions/miette.svg?maxAge=2592000
   :target: https://pypi.python.org/pypi/miette
