VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Document_Open()
    Dim ur  As UndoRecord
    Dim objField  As Field

    On Error Resume Next
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "更新所有序号"
    'Update fields
    For Each objField In ActiveDocument.Fields
       If objField.Type = wdFieldSequence Then
          objField.Update
       End If
    Next
    ur.EndCustomRecord
    'If no other change is made (except for updating the fields), allow closing without saving
    ActiveDocument.Saved = True
    'Check for standard page setup
    RestorePageSetup
End Sub
