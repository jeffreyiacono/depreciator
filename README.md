# depreciator.bas #

`depreciate` worksheet function facilitates the depreciation of a set of capital expenditures against a given depreciation schedule.

## Basic Usage ##

Import the module and use the `=depreciate` worksheet function in any cell.

The function takes two parameters:

1. `ordered_capital_expenditures`: a chronologically ordered range of capital expenditures
2. `ordered_depreciation_schedule`: the chronologically ordered depreciation schedule

Requirements:

- `ordered_capital_expenditures` must be an n x 1 or 1 x n vector
- `ordered_depreciation_schedule` must be an m x 1 or 1 x m vector

As an example, pretend it is Mar 11 and we have the following cap ex and depreciation schedule:

<pre>
   Capital Expenditures       Depreciation Schedule
   ---------------------      ---------------------
   |   Date  |  Cap Ex |      |  Month  |  Depr % |
   ---------------------      ---------------------
   |  Jan 11 |  $1000  |      |    1    |   49%   |
   |  Feb 11 |  $2000  |      |    2    |   13%   |
   |  Mar 11 |  $3000  |      |    3    |    7%   |
   ---------------------      |   ...   |   ...   |
                              |    m    |    Y%   |
                              ---------------------
</pre>

To find the depreciated cap ex, we'd pass the following into `=depreciate(ordered_capital_expenditures, ordered_depreciation_schedule)`:
_note: {} denotes range selection_

- `ordered_capital_expenditures`  = {$1000, $2000, $3000}     (vector size = 3)
- `ordered_depreciation_schedule` = {49%, 13%, 7%, ... , %Y}  (vector size = m)

The function will depreciate each month's cap ex by the corresponding entry in the depreciation schedule (_ex._ ($3,000 * 49% + $2,000 * 13% + $1,000 * 7%)).

## Sample App ##

You can view a sample application in the `/sample` directory.

## Excel Add-in ##

Excel Add-in is available in the `/addin` directory for easy distribution and installation.

##MIT License

Copyright (c) 2012 ElegantBuild, LLC, http://elegantbuild.com/

Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
