# vba-snippets
My favorite Visual Basic snippets

## Last row

Find the last row of a range (column)

``` lastrow = Range("A" & Rows.Count).End(xlUp).Row ```

## Fast range copy

Do not use paste, copy with destination is faster

``` Sheets(1).Range("A1:A10").Copy Destination:=Sheets(2).Range("A1") ```

## Fastest range copy

It's even faster if you just set the value of a range as the value of another range

``` Sheets(2).Range("A1:A10").Value = Sheets(1).Range("A1:A10").Value ```

## Copy filtered rows

With header:

``` ActiveSheet.AutoFilter.Range.Copy ```

Without the header:

``` ActiveSheet.AutoFilter.Range.Offset(1, 0).Copy ```



