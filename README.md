# ConsoleColorsTool

A little tool to display or remove the ConsoleColor settings from a .lnk file.

Note that the code is kind of shabby, but sufficient for a throw-away command line tool.
Just beware of that, should you consider integrating it into something bigger or more
worth while (especially the memory handling in the P/Invoke code would need some more
work in the face of exceptions).
