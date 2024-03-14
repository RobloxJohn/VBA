
' Sub procedure to list the Fibonacci series for all values below 1,000
Sub Fibonacci()
Dim i As Integer   ' counter for the position in the series
Dim iFib As Integer   ' stores the current value in the series
Dim iFib_Next As Integer   ' stores the next value in the series
Dim iStep As Integer   ' stores the next step size
' Initialise the variables i and iFib_Next
i = 1
iFib_Next = 0
' Do While loop to be executed as long as the value of the
' current Fibonacci number exceeds 1000
Do While iFib_Next < 1000
If i = 1 Then
' Special case for the first entry of the series
iStep = 1
iFib = 0
Else
' Store the next step size, before overwriting the
' current entry of the series
iStep = iFib
iFib = iFib_Next
End If
' Print the current Fibonacci value to column A of the
' current Worksheet
Cells(i, 1).Value = iFib
' Calculate the next value in the series and increment
' the position marker by 1
iFib_Next = iFib + iStep
i = i + 1
Loop
End Sub
