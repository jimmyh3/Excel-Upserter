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
