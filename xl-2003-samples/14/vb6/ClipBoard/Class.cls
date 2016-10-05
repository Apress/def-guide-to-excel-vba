VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClipBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' clear clipboard
Sub Clear()
  On Error Resume Next
  ClipBoard.Clear
End Sub

' read text from clipboard
Function GetText$(Optional format)
  On Error Resume Next
  If IsMissing(format) Then
    GetText = ClipBoard.GetText(1)
  Else
    GetText = ClipBoard.GetText(format)
  End If
End Function

' insert text into clipboard
Sub SetText(txt$, Optional format)
  On Error Resume Next
  If IsMissing(format) Then
    ClipBoard.SetText txt
  Else
    ClipBoard.SetText txt, format
  End If
End Sub

' test if clipboard contains data in a certain format
Function GetFormat(format)
  On Error Resume Next
  GetFormat = ClipBoard.GetFormat(format)
End Function

