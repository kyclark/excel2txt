# NAME

excel2txt - convert Excel data to delimited text files

# SYNOPSIS

    excel2txt [options] File1.xls [File2.xls ...]

Options:

    -d|--ofs                Output field delimiter (default is Tab)
    -f|--output-format      "txt," "html," "xml" or "yaml" (defaults to "txt")
    -n|--normalize-headers  Normalize column headers (see below)
    -o|--out-dir            Where to place output file (defaults to CWD)
    -q|--quiet              Do not print any status messages

    --help                  Show brief help and exit
    --man                   Show full documentation
    --version               Show version and exit

# DESCRIPTION

For each worksheet within an Excel spreadsheet, creates a text file.
By default, the output files will be plaint text files using a Tab for 
the delimiter.  Use the "-d" switch to specify a different delimiter
such as a comma.  You may also choose to create an HTML &gt;table&lt;,
an XML file, or a YAML dump using the "-f" option.

The output file names will be normalized such that they will consist
of only lowercase letters with spaces replaced by underscores and
non-alphabetic characters deleted.  The "-n" option will also apply this
transformation the column headers.  If there is only one worksheet in
an spreadsheet, then the output file will simply be the spreadsheet's
name;  if there is more than one worksheet, then a separate output
file will be created using the spreadsheet's name plus the worksheet's
name.  In any event where the default output file exists and is of a
non-zero size, then a "-1" (or "-2," etc.) will be added until a file
name is found that is not in use.

By default, progress messages are printed.  If you do not wish to see
these, use the "-q" flag.

# SEE ALSO

Spreadsheet::ParseExcel, http://code.google.com/p/perl-excel2txt/.

# AUTHOR

Ken Youens-Clark <kclark@cpan.org>.

# COPYRIGHT

Copyright (c) 2005-10 Ken Youens-Clark

This library is free software;  you can redistribute it and/or modify 
it under the same terms as Perl itself.
