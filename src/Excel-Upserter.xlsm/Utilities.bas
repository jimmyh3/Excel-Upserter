Attribute VB_Name = "Utilities"
Option Explicit

Function GetTableByName(tableName As String) As ListObject
    Dim sheet As Worksheet
    Dim table As ListObject
    
    For Each sheet In ActiveWorkbook.Worksheets
        For Each table In sheet.ListObjects
            If table.Name = tableName Then
                Set GetTableByName = table
            End If
        Next table
    Next sheet
    
End Function

Public Function GetArrayLen(arr As Variant) As Integer
    GetArrayLen = UBound(arr) - LBound(arr) + 1
End Function

Public Function Resize2DArray(multiArray As Variant, newTotalRows As Integer, newTotalCols As Integer) As Variant
    Dim result() As Variant
    ReDim result(LBound(multiArray, 1) To newTotalRows, LBound(multiArray, 2) To newTotalCols)
    Dim upperBRow As Integer
    Dim upperBCol As Integer
    Dim ridx As Integer, cidx As Integer
    upperBRow = IIf(UBound(multiArray, 1) <= newTotalRows, UBound(multiArray, 1), newTotalRows)
    upperBCol = IIf(UBound(multiArray, 2) <= newTotalCols, UBound(multiArray, 2), newTotalCols)
    
    For ridx = LBound(multiArray, 1) To upperBRow
        For cidx = LBound(multiArray, 2) To upperBCol
            result(ridx, cidx) = multiArray(ridx, cidx)
        Next cidx
    Next ridx
    
    Resize2DArray = result
End Function

Public Function RangeToArray(rng As Range) As Variant
    Dim result(): ReDim result(rng.Rows.Count, rng.Columns.Count)
    Dim row As Range
    
    For Each row In rng.Rows
        Dim cell As Range
        For Each cell In row.Cells
            'Debug.Print "Row count: " & rng.Rows.Count & "; cell address: " & cell.Address
            'Debug.Print "Row num: " & cell.row - 2 & "; cell.col: " & cell.Column
            result(cell.row - 2, cell.Column - 1) = cell.Value
        Next cell
    Next
    
    RangeToArray = result
End Function

