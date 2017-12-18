# vba-snippets
My favorite Visual Basic snippets

## Last row

Find the last row of a range (column):

    lastrow = Range("A" & Rows.Count).End(xlUp).Row

## Fast range copy

Do not use paste, copy with destination is faster:

    Sheets(1).Range("A1:A10").Copy Destination:=Sheets(2).Range("A1")

It's even faster if you just set the value of a range as the value of another range:

    Sheets(2).Range("A1:A10").Value = Sheets(1).Range("A1:A10").Value

## Copy filtered rows

With header:

    ActiveSheet.AutoFilter.Range.Copy

Without the header:

    ActiveSheet.AutoFilter.Range.Offset(1, 0).Copy

## Faster macro without updating

Turn off updating (use with caution):

    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .DisplayAlerts = False
        .AskToUpdateLinks = False
        .DisplayStatusBar = False
    End With

Don't forget to turn it on again:

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .DisplayAlerts = True
        .AskToUpdateLinks = True
        .DisplayStatusBar = True
    End With


## Remove autofilter

Check if there's an autofilter, if yes, remove it:

    If ActiveSheet.AutoFilterMode Then ActiveSheet.ShowAllData

## Reset text to column delimiter

    Dim rngEmptyCell As Range

    On Error Resume Next
        Set rngEmptyCell = ActiveSheet.Cells.SpecialCells(xlCellTypeBlanks).Cells(1, 1)
        With rngEmptyCell
            .Value = "ABC"
            .TextToColumns Destination:=rngEmptyCell, _
                DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
                ConsecutiveDelimiter:=False, Tab:=True, _
                Semicolon:=False, Comma:=False, Space:=False, _
                Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
            .Clear
        End With
    On Error GoTo 0



