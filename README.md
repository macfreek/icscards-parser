Convert bank statements from ICS cards from PDF to CSV, TSV or JSON.

ICS is a Dutch credit card issuer (icscards.nl) that allows customers to
download bank statements in PDF, not in computer readable formats. 
This Python script translates PDF bank statements to CSV transactions.

This script relies on the exact layout of the ICS bank statement.
There are checks in place to ensure that it will work on minor variations,
or report when that fails.
It has been succesfully tested on bank statements from over 10 years (2016-2026).

## Requirements

* [pdfplumber](https://pypi.org/project/pdfplumber/)

## Further Information

Homepage: https://www.github.com/macfreek/icscards-parser/

This script is copyright 2025-2026 Freek Dijkstra
License: MIT-license

In short: You can use this script in any way you like, but there is no support or warranty.
I appreciate that you give credit if you use (large parts of) this script.

## Usage

Call the script with the PDF you like to convert:

    icscards-pdf-to-csv.py rekeningoverzicht_2026-01.pdf

This will produce a file `rekeningoverzicht_2026-01.csv`.

To change the output to TSV or JSON, change the `DEFAULT_OUTPUT` global variabe in the script (on line 44). It is not yet possible to change this with a command-line argument.

To parse all files in a directory, give a directory as an argument:

    icscards-pdf-to-csv.py /path/to/rekeningen/

This will attempt to parse all PDF files with a year and month in the name (`*YYYY-MM*.PDF`). This prevents the script from accidentally parsing other PDF files. Subdirectories are not parsed, and it will skip a PDF if the output file already exists.

To save output files in a different directory, change the `DESTINATION_DIR` global variabe in the script (on line 45).

Feel free to leave any feedback at the homepage on github.
