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

Dim oldTable As ListObject
Dim updTable As ListObject 'Updated Table'
Dim matchCols() As String

Public Function Run(oldTbl As String, updTbl As String, mtchCols() As String)
    Set oldTable = GetTableByName(oldTbl)
    Set updTable = GetTableByName(updTbl)
    matchCols = mtchCols
    
    If IsInputValid Then
        Upsert
    End If
End Function

Private Function IsInputValid() As Boolean
    Dim result As Boolean
    Dim message As String
    result = True
    
    Rem SUGGESTION: use UserForm to obtain multiple user input at once
    Rem SUGGESTION: allow user to decide if initial row are column headers.
    Rem SUGGESTION: allow user to check if they want items sorted (helps computation)
    
    Rem Determine if the match columns exist in both old table and update table.
    
    IsInputValid = result
End Function

Rem Output result into new sheet (do not replace the original selections
Private Function Upsert() As Range
    Dim newWS As Worksheet
    Rem Set newWS = Sheets.Add
    
    '1. Ensure the selected third column can be found in the 2nd table.'
    '2. Scan 2nd table for any new columns and add that column to old table'
    '3. '
End Function
