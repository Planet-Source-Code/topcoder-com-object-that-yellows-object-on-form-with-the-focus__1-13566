VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChgCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author Robert Jackson
'Email Rjackson1_99@yahoo.com
'Code Makes it Easier for Users to remember where they were


'we have to declare as object so the code can be used with combo, checkboxes, etc., if you
'know object ahead of time you can change to that type of Object
Dim Gcurrent As Object
Dim Gpast As Object


Option Explicit

Public Sub ChangeCol()

If Gcurrent Is Nothing Then
        Set Gcurrent = Screen.ActiveControl
Else
    Set Gpast = Gcurrent
    Gpast.BackColor = vbWhite
    Set Gcurrent = Screen.ActiveControl
    Gcurrent.BackColor = vbYellow
End If
End Sub

