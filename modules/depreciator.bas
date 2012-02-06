Attribute VB_Name = "depreciator"
Option Explicit
' #depreciate facilitates the depreciation of a set of capitial expenditures against a given depreciation schedule.
' The function takes two parameters:
'
'   1. ordered_capital_expenditures: a chronologically ordered range of capital expenditures
'   2. ordered_depreciation_schedule: the chronologically ordered depreciation schedule
'
' Requirements: ordered_capital_expenditures must be an n x 1 or 1 x n vector
'               ordered_depreciation_schedule must be an m x 1 or 1 x m vector
'
' As an example, if we have the following:
'
'    Capital Expenditures       Depreciation Schedule
'    ---------------------      ---------------------
'    |   Date  |  Cap Ex |      |  Month  |  Depr % |
'    ---------------------      ---------------------
'    |  Jan 11 |  $1000  |      |    1    |   49%   |
'    |  Feb 11 |  $2000  |      |    2    |   13%   |
'    |  Mar 11 |  $3000  |      |    3    |    7%   |
'    ---------------------      |   ...   |   ...   |
'                               |    m    |    Y%   |
'                               ---------------------
'
' Then, we pass the following the parameters ordered as ({} denotes range selection):
'
'   => ordered_capital_expenditures  = {$1000, $2000, $3000}     (vector size = 3)
'   => ordered_depreciation_schedule = {49%, 13%, 7%, ... , %Y}  (vector size = m)
'
' Also, if the cap ex entries exceed the depreciation schedule, we mark depreciation expense to zero as they will not be eligible for depreciation anymore
'
Public Function depreciate(ordered_capital_expenditures As range, ordered_depreciation_schedule As range) As Variant
  On Error GoTo failure:
  ' ensure that ordered_capital_expenditures is a vector (n x 1 or 1 x n)
  If Not is_vector(ordered_capital_expenditures) Then
    depreciate = not_vector_error_message("ordered capital expenditures parameter", ordered_capital_expenditures)
    Exit Function
  End If
  
  ' ensure that ordered_depreciation_schedule is a vector (m x 1 or 1 x m)
  If Not is_vector(ordered_depreciation_schedule) Then
    depreciate = not_vector_error_message("ordered depreciation schedule parameter", ordered_depreciation_schedule)
    Exit Function
  End If
    
  Dim i       As Integer
  Dim i_limit As Integer
  Dim sum     As Double
  
  ' find iteration limit: if we have more ordered capital expenditures than the depreciation schedule covers, exclude them.
  ' why? because any cap ex that is beyond the full depreciation schedule should have already been completely depreciated.
  ' this does assume that full depreciation schedule is being supplied and that it sums to 100%
  i_limit = IIf(ordered_capital_expenditures.Count < ordered_depreciation_schedule.Count, ordered_capital_expenditures.Count, ordered_depreciation_schedule.Count)
  ' initiate sum to zero
  sum = 0
  ' iterate over appropriate capital expenditures, multiplying each by its corresponding depreciation schedule entry
  ' aggregate the sum each
  For i = 1 To i_limit
    sum = sum + ordered_capital_expenditures(ordered_capital_expenditures.Count - i + 1) * ordered_depreciation_schedule(i)
  Next i
  ' assign sum as return value & exit function
  depreciate = sum
  Exit Function
  
failure:
  ' on error, return #N/A
  depreciate = CVErr(xlErrNA)
End Function
' returns string formatted as "{row count} x {column count}"
Private Function stringified_dimensions(r As range) As String
  stringified_dimensions = r.Rows.Count & " x " & r.Columns.Count
End Function
' returns true if the passed range is an n x 1 or 1 x n vector
' false otherwise
Private Function is_vector(r As range) As Boolean
  is_vector = (r.Rows.Count = 1 Or r.Columns.Count = 1)
End Function
' convenience helper method to return common error message reporting that the range is not a vector
' must supply the stringified parameter name and the parameter (range) itself
' returned message if the format of:
'   "Error: parameter_name must be an n x 1 or 1 x n range. Currently: {range rows count} x {range columns count}
Private Function not_vector_error_message(parameter_name As String, rng As range) As String
  not_vector_error_message = "Error: " & parameter_name & " must be an n x 1 or 1 x n range. Currently: " & stringified_dimensions(rng)
End Function
