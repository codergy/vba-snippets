# VBA snippets
My favorite Visual Basic snippets

## Last row

Find the last row of a range (column):

    lastrow = Range("A" & Rows.Count).End(xlUp).Row

## Fast range copy

Do not use ".Paste", ".Copy" with "Destination" is much faster:

    Sheets(1).Range("A1:A10").Copy Destination:=Sheets(2).Range("A1")

It's even faster if you just set the value of a range as the value of another range:

    Sheets(2).Range("A1:A10").Value = Sheets(1).Range("A1:A10").Value

If you want to copy formulas:

    Sheets(2).Range("A1:A10").Formula = Sheets(1).Range("A1:A10").Formula

If you have formulas in a range, but you need their value only, do not use copy and paste values. Here's a fast method:

    With Sheets(2).Range("A1:A10")
        .Value = .Value
    End With

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

Check if there's an autofilter. If yes, remove it:

    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilterMode = False

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

## Read data from another workbook without opening it

Opening files can be slow, let's read data from closed files:

    p = "C:\folder\" 'source file's folder
    f = "filename.xls" 'source file
    s = "Sheet1" 'source sheet
    a = "C10" 'first cell that should be copied from source file

    'copy data to range A1:Z1000 from C:\folder\filename.xls, Sheet1, range C10:AB1010
    'then turn the formulas into values
    With Sheets(1).Range("A1:Z1000")
        .Formula = "='" & p & "[" & f & "]" & s & "'!" & a
        .Value = .Value
    End With
 
## Trim a whole range

Fast and easy to read trim function.

    With Range("A1:A100")
        .Value = Application.Trim(.Value)
    End With

## Change default folder when opening file

Browsing and opening another Excel file: it's more comfortable if the default folder changes to the one where the macro file is.

    ChDrive Left(ActiveWorkbook.Path, 2)
    ChDir ActiveWorkbook.Path

## Check if sheet (name) exists

Call the function (returns boolean value):
    
    sheetExists("Sheet3")

The function:
    
    Function sheetExists(sheetName As String) As Boolean

        Dim rng As Range

        On Error Resume Next
        Set rng = Sheets(sheetName).Range("A1")
        On Error GoTo 0

        sheetExists = Not (rng Is Nothing)

    End Function

## Check if file or folder exists

**Check if file exists**

Call the function (returns boolean value):

    fileExists(ThisWorkbook.Path, ThisWorkbook.Name)

The function:

    Function fileExists(fpath As String, fname As String) As Boolean

        fileExists = (Dir(fpath & "\" & fname) <> vbNullString)

    End Function

**Check if folder exists**

Call the function (returns boolean value):

    folderExists(ThisWorkbook.Path)

The function:

    Function folderExists(folderName As String) As Boolean

        folderExists = (Dir(folderName, vbDirectory) <> vbNullString)

    End Function

