## depreciator ##

`depreciate` worksheet function facilitates the depreciation of a set of capital expenditures against a given depreciation schedule.

## Basic Usage ##

Import the module and use the `=depreciate` worksheet function in any cell.

The function takes two parameters:

1. `ordered_capital_expenditures`: a chronologically ordered range of capital expenditures
2. `ordered_depreciation_schedule`: the chronologically ordered depreciation schedule

Requirements:

- `ordered_capital_expenditures` must be an n x 1 or 1 x n vector
- `ordered_depreciation_schedule` must be an m x 1 or 1 x m vector

As an example, if we have the following:

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

Then, we pass the following the parameters ordered as ({} denotes range selection):

- `ordered_capital_expenditures`  = {$1000, $2000, $3000}     (vector size = 3)
- `ordered_depreciation_schedule` = {49%, 13%, 7%, ... , %Y}  (vector size = m)

Also, if the cap ex entries exceed the depreciation schedule, we mark the depreciation expense to zero as it will not be eligible for depreciation anymore.

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
