VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_Upserter 
   Caption         =   "Upserter"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UF_Upserter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_Upserter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Btn_Upsert_Click()
    UPSERT_TABLES.Run ComboBox_OldTable.Value, ComboBox_UpdatedTable.Value, ListBox_MatchColumns.Value
    Unload Me
End Sub

Private Sub ComboBox_OldTable_Change()
    ListBox_MatchColumns.Clear
    
    If Not IsEmpty(ComboBox_OldTable.Value) Then
        Dim header As Range
        Dim table As ListObject
        Set table = GetTableByName(ComboBox_OldTable.Value)
        
        For Each header In table.HeaderRowRange
            ListBox_MatchColumns.AddItem header.Value
        Next header
    End If
End Sub

Private Sub ListBox_MatchColumns_Click()

End Sub

Private Sub UserForm_Initialize()
    InitializeTableListing
End Sub

Private Sub InitializeTableListing()
    Dim sheet As Worksheet
    Dim table As ListObject 'ListObject are tables in Excel'
    
    For Each sheet In ActiveWorkbook.Worksheets
        For Each table In sheet.ListObjects
            ComboBox_OldTable.AddItem table
            ComboBox_UpdatedTable.AddItem table
        Next table
    Next sheet
End Sub
