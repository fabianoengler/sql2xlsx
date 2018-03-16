# sql2xlsx

Easily generate formatted Excel/OpenXML/XLSX files from SQL queries to MySQL. 

The generated spreadsheet will have the columns names from the query on
the first line of the working sheet. The working sheet will have the first
line frozen and will have auto filter enabled for all columns.

Cells' types will match the returned types from MySQL, eg., numbers will be
number, dates will be dates, etc.

Floats and Decimals will be formatted as numbers with two decimal digits and
with decimal and thousands separators. Other types will use the standard
formatting for each type.

Columns' width will be resized to fit at least 90% of the cells in the column.


## How to use

Suppose you have a MySQL select query in a file called `myquery.sql` and want
to export the results from executing this query to an Excel .xlsx file.

You can simply run:

    $ ./sql2xlsx.py myquery.sql myreport.xlsx

**And that is it.**

_Note: Output file will be overwritten without confirmation if it already
exists._

The second argument is optional, if omitted, the output file will be created
on the current directory with a timestamp and .xlsx extension. For example:
`sql2xlsx /home/user/reports/monthly-sales.sql` will output to
`./2018-03-16_160450_monthly-sales.xlsx` . 


## Quick Start

Basic steps are:

1. Download this script (clone this repo for example)
2. Install dependencies (connector for mysql and openpyxl)
3. Setup db config (username, password, db name, hostname)
4. Run the script =)

All the steps above from unix shell command line:

    $ git clone https://github.com/fabianoengler/sql2xlsx.git
    $ cd sql2xlsx
    $ pip install -r requirements.txt
    $ pip install lxml  # optional, for better performance
    $ cp config_sample.py config.py
    $ vim config.py  # or whatever editor you prefer

And you are done, just run the script now.

If you don't have an SQL file lurking around already, you can quickly
create one for testing:

    $ echo 'SELECT * FROM users' > all-users.sql
    $ ./sql2xlsx.py all-users.sql all-users.xlsx


## Dependencies and Python Versions

This script was tested with
- Python 3.4 and 3.6
- mysql-connector-python-8.0.6
- openpyxl-2.5.1
- lxml-4.2.0

It may work with Python 3.3 as well (not tested though).

It may work with other mysql-connectors (not tested though).


## Misc Info


I had a lot of SQL files for MySQL around that I used to copy and paste to
phpMyAdmin and export the results when I needed a quick report or a quick dump
of a table.

So I decided to write a helper script to save me a few minutes when I needed
to export those results to a spreadsheet and apply some filtering and sorting.

Hence, this is very simple and hacky script done in a hurry in a rainy night
that I decided to share, so don't expect much: It doesn't have unittest, it
doesn't use argparse, it's source-code is not commented, it doesn't use the
best coding practices ever, it doesn't have a lot of error handling, etc., etc.

Again: it is a very simple script done in a few spare hours, but it does
the job for me. Hope you find it useful as well.


## Customizing the XLSX output

The script is extremely simple. It has one class called `MySql2Xlsx` that
drives all the process of executing the query and writing the results to a
spreadsheet.

If called from command line, the script will instantiate an object of this
class and simply call `generate_report()`.

`generate_report()` in the other hand will simply call many helper methods
in sequence. If you want to customize anything, it should be very easy to
subclass `MySql2Xlsx` and overwrite any methods you want.

Some good candidates for overwriting are:

- `make_final_adjustments()`
- `resize_columns()`
- `format_numbers()`
- `format_column_names()`


## Hidden feature: parameterized queries

The queries can actually have parameters, using python format style, like:

```SQL
SELECT
    first_name, last_name, hire_date
FROM employees
WHERE hire_date
    BETWEEN %(start_date)s AND %(end_date)s
```


But the parameters are not yet supported to be passed from the command line.

If you want to used this, for now you will have to instantiate the class
for yourself. A complete example on how to do that:


```python
#!/bin/env python3

from sql2xlsx import MySql2Xlsx
from config import mysql_config 
import sys
import datetime

query = open(sys.argv[1], 'rt').read()
out_fname = sys.argv[2]

params = {
    'start_date': datetime.date(2017, 1, 1),
    'end_date': datetime.date(2017, 6, 1)
}

mysql2xlsx = MySql2Xlsx(mysql_config, query, out_fname, params)
mysql2xlsx.generate_report()
```

## To Do / Next Steps / Roadmap / Ideas

This script was hacked in a few hours, without proper development practices.

Below is a list of some things I would like to add/change one day (in no
particular order). Or ideas for you to contribute if you feel like it:

- [x] Convert procedural steps to methods in a class.
- [x] Make the class easy to be used/subclassed by importing scripts.
- [x] Accept parameterized queries.
- [x] Implement some heuristic for column width resizing.
- [x] Use tempfile for intermediate file if output file name not defined.
- [x] Add timestamp to derived output name.
- [x] Make derived output file path to be current working dir.
- [ ] Add tests (unittest or BDD?).
- [ ] Use argparse for command line options handling.
- [ ] Create a decent --help.
- [ ] Document the source code.
- [ ] Make a distributable package (installable via pip).
- [ ] Create a man page.
- [ ] Make internal verbosity level configurable by CLI (eg. via
      multiple -v options).
- [ ] Add parameterized queries to CLI.
- [ ] Improve heuristic for column width resizing and make it easier to
      overwrite/customize.
- [ ] Make it easier to customize common formatting, such as default
      font and size.
- [ ] Make the column resizing algorithm consider the font size.
- [ ] Make the column resizing algorithm handle multiline cell contents.
- [ ] Replace custom hacky printing for logging module.
- [ ] Make use of config file optional (credentials via CLI or env vars).
- [ ] Check if output file exists and add a flag to force overwrite (like -f)
- [ ] Add option to output CSV instead of XLSX?


**If you got any other ideas, feel free to reach me out or fork and submit a
pull request!**


