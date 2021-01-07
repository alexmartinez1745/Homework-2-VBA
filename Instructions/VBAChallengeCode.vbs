Sub Stocks()
'Set variable to run through all worksheets
    Dim WS As Worksheet
'Loop through all worksheets
    For Each WS In Worksheets
'Name columns
        WS.Range("I1").Value = "Ticker"
        WS.Range("J1").Value = "Yearly Change"
        WS.Range("K1").Value = "Percent Change"
        WS.Range("L1").Value = "Total Stock Volume"
'Start with creating a variable for ticker/summary symbols
        Dim ticker As String
        Dim Summary As Long
        Summary = 2
'Creating variable for yearly change
        Dim YearOpen As Double
         YearOpen = 0
        Dim YearClose As Double
         YearClose = 0
        Dim YearChange As Double
         YearChange = 0
'Create variable for Percentage change
        Dim Percentage As Double
         Percentage = 0
'Create variable for Volume of each stock
        Dim Vol As Double
         Vol = 0
'Loop through the sheet with last row shortcut
        Lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row
'Set begin price for when stock opens
        YearOpen = WS.Cells(2, 3).Value
'Loop through worksheet
        For i = 2 To Lastrow
'Find when ticker changes and set up in designated column
            If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
                ticker = WS.Cells(i, 1).Value
'Finding yearly change for each ticker category
                YearClose = WS.Cells(i, 6).Value
                YearChange = YearClose - YearOpen
'Find percentage change
                If YearOpen <> 0 Then
                    Percentage = (YearChange / YearOpen) * 100
                End If
'Calculate total stock volume for each category
                Vol = Vol + WS.Cells(i, 7).Value
'Print to Summary Table
                WS.Range("I" & Summary).Value = ticker
                WS.Range("J" & Summary).Value = YearChange
                WS.Range("K" & Summary).Value = (Str(Percentage) & "%")
                WS.Range("L" & Summary).Value = Vol
'Highlighting positive and negative change for yearchage
                    If (YearChange < 0) Then
                        WS.Range("J" & Summary).Interior.ColorIndex = 3
                    ElseIf (YearChange > 0) Then
                        WS.Range("J" & Summary).Interior.ColorIndex = 4
                    End If
'Add to summary row to go through each row change
                Summary = Summary + 1
'Reset Variables
                YearOpen = WS.Cells(i + 1, 3)
                YearClose = 0
                YearChange = 0
                Vol = 0
                Percentage = 0
            Else
                Vol = Vol + WS.Cells(i, 7).Value
            End If
        Next i
    Next WS
End Sub