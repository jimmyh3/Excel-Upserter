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
Dim result_dataBodyArray() As Variant
Dim keyCols() As String

Public Function Run(oldTbl As String, updTbl As String, keyColumns() As String)
    Application.ScreenUpdating = False
    Dim newWS As Worksheet
    Set newWS = Sheets.Add
    Set toUpsertTable = GetTableByName(oldTbl)
    Set upserterTable = GetTableByName(updTbl)
    keyCols = keyColumns
    
    Call toUpsertTable.Range.Copy(Destination:=newWS.Range("A1"))
    Set toUpsertTable = newWS.ListObjects(1)
    
    If hasKeyColumns(upserterTable) Then
        Call addMissingTableColumns(toUpsertTable, upserterTable)
        result_dataBodyArray = toUpsertTable.DataBodyRange.Value
        Call Upsert(toUpsertTable, upserterTable)
        
        'Resize table to accomodate resulting array otherwise it's truncated'
        toUpsertTable.Resize toUpsertTable.Range.Resize(UBound(result_dataBodyArray, 1) + 1, UBound(result_dataBodyArray, 2))
        toUpsertTable.DataBodyRange.Value = result_dataBodyArray
    Else
        Application.DisplayAlerts = False
        newWS.ListObjects(1).Delete
        newWS.Delete
        Application.DisplayAlerts = True
    End If
    
    Application.ScreenUpdating = True
End Function

Private Function hasKeyColumns(table_updater As ListObject) As Boolean
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
        For uidx = 1 To table_updater.HeaderRowRange.Count
            If mcol = table_updater.HeaderRowRange(uidx) Then
                hasNeededCols = True
                Exit For
            End If
        Next uidx
        
        If hasNeededCols = False Then
            MsgBox "Error: 2nd table_updater needs to have the selected key columns."
            hasKeyColumns = False
            Exit Function
        End If
    Next mcol
    
    'Every test case so far will exit function with result'
    hasKeyColumns = True
    
End Function

Private Sub addMissingTableColumns(table_result As ListObject, table_toReference As ListObject)
    Dim hasCol As Boolean, uidx As Integer, oldCol As Variant
    For uidx = 1 To table_toReference.ListColumns.Count
        hasCol = False
        
        For Each oldCol In table_result.ListColumns
            If table_toReference.ListColumns(uidx) = oldCol Then
                hasCol = True
                Exit For
            End If
        Next oldCol
        
        If hasCol = False Then
            Dim addIdx As Integer
            addIdx = table_result.ListColumns.Count + 1
            table_result.ListColumns.Add (addIdx)
            table_result.ListColumns(addIdx).Name = table_toReference.ListColumns(uidx)
        End If
    Next uidx
End Sub

Rem Output result into new sheet (do not replace the original selections
Private Function Upsert(table_toUpdate As ListObject, table_updater As ListObject) As Range

    'Thinking too long on this; just write what you wish it can do..'
    '1. Loop through every row of upserterTable'
    '2. For each upserterTable row, by cell, find first instance of toUpsertTable value that matches'
    '3. Create Inner loop for #2. to then find subsequent key column values by matching column names'
    Dim rcount As Integer
    
    'Iterate each row in upserterTable'
    For rcount = 1 To table_updater.DataBodyRange.Rows.Count
        'Check if current row should be updated or inserted'
        Dim result As Integer: result = isUpdateRow(table_toUpdate, table_updater, rcount)
        If result > 0 Then
            Call updateRow(table_toUpdate, table_updater, result, rcount)
        Else
            Call insertRow(table_toUpdate, table_updater, rcount)
        End If
    Next rcount
    
End Function

Private Function isUpdateRow(table_toUpdate As ListObject, table_updater As ListObject, rcount As Integer) As Long
    Dim isUpdate As Boolean: isUpdate = True
    'keyCell = cell in upserterTable that acts as initial value to find in cross referencing'
    Dim keyCell As Range: Set keyCell = table_updater.DataBodyRange.Cells(rcount, upserterTable.ListColumns(keyCols(0)).Index)
    'colToFindIn = column in toUpsertTable to find the keyCell value in'
    Dim colToFindIn As Range: Set colToFindIn = table_toUpdate.ListColumns(keyCols(0)).Range
    'firstFind = 1st cell in toUpsertTable under the same column name as keyCell [over at upserterTable] with the same value'
    Dim firstFind As Range: Set firstFind = colToFindIn.Find(keyCell.Value, LookIn:=xlValues)
    'loopFind is needed to continue Range.FindNext; firstFind is kept to ensure we don't continue searching forever'
    Dim loopFind As Range: Set loopFind = firstFind
    Dim result As Integer: result = -1
    
    If Not firstFind Is Nothing Then
        Do
            Dim hasAllKeyMatches As Boolean: hasAllKeyMatches = True
            Dim keyColIdx As Integer
            For keyColIdx = 1 To UBound(keyCols)
                Dim updater_colnum As Integer: updater_colnum = table_updater.ListColumns(keyCols(keyColIdx)).Index
                Dim updater_rownum As Integer: updater_rownum = rcount
                Dim updater_cell As Range: Set updater_cell = table_updater.DataBodyRange.Cells(updater_rownum, updater_colnum)
                Dim toUpdate_colnum As Integer: toUpdate_colnum = table_toUpdate.ListColumns(keyCols(keyColIdx)).Index
                Dim toUpdate_rownum As Integer: toUpdate_rownum = loopFind.row
                Dim toUpdate_cell As Range: Set toUpdate_cell = table_toUpdate.Range.Cells(toUpdate_rownum, toUpdate_colnum)
                
                If toUpdate_cell.Value <> updater_cell.Value Then
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
    Else
        'Debug.Print "Cannot find initial key cell: " & keyCell.Value & "; colToFind is: " & colToFindIn.Address
    End If
    
    
End Function
    
Private Function updateRow(table_toUpdate As ListObject, table_updater As ListObject, toUpdate_rownum As Integer, updater_rownum As Integer) As Integer

    Dim toUpdate_cell As Range
    Dim updater_cell As Range
    
    Dim colnameVar As Variant
    For Each colnameVar In table_updater.ListColumns
        Dim colname As String: colname = CStr(colnameVar)
        Set updater_cell = table_updater.DataBodyRange.Cells(updater_rownum, table_updater.ListColumns(colname).Index)
        result_dataBodyArray(toUpdate_rownum, table_toUpdate.ListColumns(colname).Index) = updater_cell.Value
    Next colnameVar
        
End Function

Private Function insertRow(table_toUpdate As ListObject, table_updater As ListObject, updater_rownum As Integer) As Integer
    result_dataBodyArray = Resize2DArray(result_dataBodyArray, UBound(result_dataBodyArray, 1) + 1, UBound(result_dataBodyArray, 2))
        
    Dim colnameVar As Variant
    For Each colnameVar In table_updater.ListColumns
        Dim colname As String: colname = CStr(colnameVar)
        result_dataBodyArray(UBound(result_dataBodyArray, 1), table_toUpdate.ListColumns(colname).Index) = table_updater.DataBodyRange.Cells(updater_rownum, table_updater.ListColumns(colname).Index).Value
    Next colnameVar
End Function


