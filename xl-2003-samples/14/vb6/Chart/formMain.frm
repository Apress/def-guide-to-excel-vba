VERSION 5.00
Begin VB.MDIForm formMain 
   BackColor       =   &H8000000C&
   Caption         =   "Show 3D chart using Excel"
   ClientHeight    =   5940
   ClientLeft      =   435
   ClientTop       =   1770
   ClientWidth     =   6690
   Icon            =   "formMain.frx":0000
   LinkTopic       =   "MDIForm1"
End
Attribute VB_Name = "formMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' vb6\chart\formMain.frm
Option Explicit
Dim olef As New formOLE

Private Sub MDIForm_Load()
  olef.Show                  'show OLEForm
End Sub
