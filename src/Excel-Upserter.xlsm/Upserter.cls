VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Upserter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim toUpsertTable As ListObject
Dim upserterTable As ListObject 'Updated Table'
Dim resTable As ListObject
Dim keyCols() As String

Public Function Run(oldTbl As String, updTbl As String, keyColumns() As String)
    Dim newWS As Worksheet
    Set newWS = Sheets.Add
    Set toUpsertTable = GetTableByName(oldTbl)
    Set upserterTable = GetTableByName(updTbl)
    keyCols = keyColumns
    
    Call toUpsertTable.Range.Copy(Destination:=newWS.Range("A1"))
    Set toUpsertTable = newWS.ListObjects(1)
    
    If IsInputValid(toUpsertTable, upserterTable) Then
        Call addMissingColsTotoUpsertTable(toUpsertTable, upserterTable)
        Call Upsert(toUpsertTable, upserterTable)
    Else
        Application.DisplayAlerts = False
        newWS.ListObjects(1).Delete
        newWS.Delete
        Application.DisplayAlerts = True
    End If
    
End Function

Private Function IsInputValid(toUpsertTable_ As ListObject, upserterTable_ As ListObject) As Boolean
    IsInputValid = hasKeyColumns(upserterTable_)
End Function

Private Function hasKeyColumns(upserterTable_ As ListObject) As Boolean
    Dim mcol As Variant
    Dim uidx As Integer
    Dim isEmptyArr As Boolean
    Dim hasNeededCols As Boolean
    isEmptyArr = True
    hasNeededCols = False
    
    On Error Resume Next
        isEmptyArr = Not IsNumeric(UBound(keyCols))
        If isEmptyArr = True Then
            MsgBox "Error: Require key column selection!"
            hasKeyColumns = False
            Exit Function
        End If
    On Error GoTo 0
        
    For Each mcol In keyCols
        hasNeededCols = False
        For uidx = 1 To upserterTable_.HeaderRowRange.Count
            If mcol = upserterTable_.HeaderRowRange(uidx) Then
                hasNeededCols = True
                Exit For
            End If
        Next uidx
        
        If hasNeededCols = False Then
            MsgBox "Error: 2nd Table needs to have the selected key columns."
            hasKeyColumns = False
            Exit Function
        End If
    Next mcol
    
    'Every test case so far will exit function with result'
    hasKeyColumns = True
    
End Function

Private Sub addMissingColsTotoUpsertTable(toUpsertTable_ As ListObject, upserterTable_ As ListObject)
    Dim hasCol As Boolean, uidx As Integer, oldCol As Variant
    For uidx = 1 To upserterTable_.ListColumns.Count
        hasCol = False
        
        For Each oldCol In toUpsertTable_.ListColumns
            If upserterTable_.ListColumns(uidx) = oldCol Then
                hasCol = True
                Exit For
            End If
        Next oldCol
        
        If hasCol = False Then
            Dim addIdx As Integer
            addIdx = toUpsertTable_.ListColumns.Count + 1
            toUpsertTable_.ListColumns.Add (addIdx)
            toUpsertTable_.ListColumns(addIdx).Name = upserterTable_.ListColumns(uidx)
        End If
    Next uidx
End Sub

Rem Output result into new sheet (do not replace the original selections
Private Function Upsert(toUpsertTable_ As ListObject, upserterTable_ As ListObject) As Range

    'Thinking too long on this; just write what you wish it can do..'
    '1. Loop through every row of upserterTable'
    '2. For each upserterTable row, by cell, find first instance of toUpsertTable value that matches'
    '3. Create Inner loop for #2. to then find subsequent key column values by matching column names'
    Dim ccount As Integer
    Dim rcount As Integer
    ccount = upserterTable.DataBodyRange.Columns.Count
    
    'Iterate each row in upserterTable'
    For rcount = 1 To upserterTable.DataBodyRange.Rows.Count
        'Check if current row should be updated or inserted'
        Dim result As Integer: result = isUpdateRow(toUpsertTable_, upserterTable_, rcount)
        If result > 0 Then
            Rem MsgBox "Is update row: " & upserterTable.ListRows(rcount).Range.Address & " for toUpdateTable at " & toUpsertTable_.ListRows(result).Range.Address
            Call updateRow(toUpsertTable_, upserterTable_, result, rcount)
        Else
            Rem MsgBox "Is NOT an update row: " & upserterTable.ListRows(rcount).Range.Address
        
        End If
    Next rcount
    
End Function

Private Function isUpdateRow(toUpsertTable_ As ListObject, upserterTable_ As ListObject, rcount As Integer) As Long
    Dim isUpdate As Boolean: isUpdate = True
    'keyCell = cell in upserterTable that acts as initial value to find in cross referencing'
    Dim keyCell As Range: Set keyCell = upserterTable_.DataBodyRange.Cells(rcount, upserterTable.ListColumns(keyCols(0)).Index)
    'colToFindIn = column in toUpsertTable to find the keyCell value in'
    Dim colToFindIn As Range: Set colToFindIn = toUpsertTable_.ListColumns(keyCols(0)).Range
    'firstFind = 1st cell in toUpsertTable under the same column name as keyCell [over at upserterTable] with the same value'
    Dim firstFind As Range: Set firstFind = colToFindIn.Find(keyCell.Value, LookIn:=xlValues)
    'loopFind is needed to continue Range.FindNext; firstFind is kept to ensure we don't continue searching forever'
    Dim loopFind As Range: Set loopFind = firstFind
    Dim result As Integer: result = -1
        
    If Not firstFind Is Nothing Then
        Rem MsgBox "KeyCell: " & keyCell & "; Found match in toUpsertTable at: " & firstFind.Address & " as " & firstFind.Value
        
        Do
            Dim hasAllKeyMatches As Boolean: hasAllKeyMatches = True
            Dim keyColIdx As Integer
            For keyColIdx = 1 To UBound(keyCols)
                Dim upsert_col As Integer: upsert_col = upserterTable_.ListColumns(keyCols(keyColIdx)).Index
                Dim upsert_row As Integer: upsert_row = rcount
                Rem MsgBox "upsert_col: " & upsert_col & "; upsert_row: " & upsert_row
                Dim upsert_cell As Range: Set upsert_cell = upserterTable_.DataBodyRange.Cells(upsert_row, upsert_col)
                Dim toUpT_colnum As Integer: toUpT_colnum = toUpsertTable_.ListColumns(keyCols(keyColIdx)).Index
                Dim toUpT_rownum As Integer: toUpT_rownum = loopFind.row
                Rem MsgBox "toUpT_colnum: " & toUpT_colnum & " toUpT_rownum: " & toUpT_rownum
                Dim toUpT_cell As Range: Set toUpT_cell = toUpsertTable_.Range.Cells(toUpT_rownum, toUpT_colnum)
                Rem MsgBox "upsert_cell: " & upsert_cell.Value & " " & upsert_cell.Address & " ; toUpt_cell: " & toUpT_cell.Value & " " & toUpT_cell.Address
                
                If toUpT_cell.Value <> upsert_cell.Value Then
                    hasAllKeyMatches = False
                    Exit For
                End If
            Next keyColIdx
            
            If hasAllKeyMatches = True Then
                result = loopFind.row - 1 'WARNING: -1 because of Range and DataBodyRange offset error; ListObject.ListColumns or ListRows refer to whole table, not just DataBodyRange'
                isUpdateRow = result
                Exit Function
            End If
            
            Set loopFind = colToFindIn.FindNext(loopFind)
        Loop While Not loopFind Is Nothing And firstFind.Address <> loopFind.Address
    End If
    
End Function
    
Private Function updateRow(toUpsertTable_ As ListObject, upserterTable_ As ListObject, toUpT_rownum As Integer, upsert_rownum As Integer) As Integer

        Dim toUpTbl_cell As Range
        Dim upsertTbl_cell As Range
        
        Dim colnameVar As Variant
        For Each colnameVar In upserterTable_.ListColumns
            Dim colname As String: colname = CStr(colnameVar)
            Set toUpTbl_cell = toUpsertTable_.DataBodyRange.Cells(toUpT_rownum, toUpsertTable_.ListColumns(colname).Index)
            Set upsertTbl_cell = upserterTable_.DataBodyRange.Cells(upsert_rownum, upserterTable_.ListColumns(colname).Index)
            Rem MsgBox "toUpTbl_cell address: " & toUpTbl_cell.Address & "; upsertTbl_cell: " & upsertTbl_cell.Address
            Rem MsgBox "Replace toUpTbl_cell: " & toUpTbl_cell.Value & " with: " & upsertTbl_cell.Value
            toUpTbl_cell.Value = upsertTbl_cell.Value
        Next colnameVar
        
End Function
