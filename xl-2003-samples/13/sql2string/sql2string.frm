VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Transform SQL code into a VB string"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtVarname 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   "sql"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtResult 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Text            =   "sql2string.frx":0000
      Top             =   2640
      Width           =   5895
   End
   Begin VB.TextBox txtSQL 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Text            =   "sql2string.frx":0020
      Top             =   840
      Width           =   5895
   End
   Begin VB.Label Label3 
      Caption         =   "Variable name:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblResult 
      Caption         =   "Resulting VB code:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2400
      Width           =   5655
   End
   Begin VB.Label lblSQL 
      Caption         =   "Paste the SQL command (using the clipboard / Strg+V)."
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const maxlines = 20

Private Sub txtVarname_Change()
  txtSQL_Change
End Sub
Private Sub txtSQL_Change()
  Dim sqllines As Variant, vblines As Variant
  Dim i&
  sqllines = Split(txtSQL, vbCrLf)
  For i = 0 To UBound(sqllines)
    If i = UBound(sqllines) And sqllines(i) = "" Then Exit For
    sqllines(i) = strng(sqllines(i))
    If i = 0 Then
      sqllines(i) = txtVarname + " = " & sqllines(i)
    ElseIf (i Mod maxlines) = 0 Then
      sqllines(i) = txtVarname + " = " + txtVarname + " + " + sqllines(i)
    Else
      sqllines(i) = String(Len(txtVarname), " ") + "   " + sqllines(i)
    End If
    If i < UBound(sqllines) And ((i + 1) Mod maxlines) <> 0 _
       And Not (i + 1 = UBound(sqllines) And sqllines(UBound(sqllines)) = "") Then
      sqllines(i) = sqllines(i) + " + _"
    End If
  Next
  txtResult = Join(sqllines, vbCrLf)
End Sub

Private Function strng(x)
  strng = """" + Replace(x, """", """""") + " """
End Function

' mark content of text boxes
Private Sub txtSQL_GotFocus()
  txtSQL.SelStart = 0
  txtSQL.SelLength = Len(txtSQL)
End Sub
Private Sub txtResult_GotFocus()
  txtResult.SelStart = 0
  txtResult.SelLength = Len(txtResult)
  Clipboard.Clear
  Clipboard.SetText txtResult
End Sub

' Resize
Private Sub Form_Resize()
  On Error Resume Next
  txtSQL.Move 0, txtSQL.Top, ScaleWidth, (ScaleHeight - txtSQL.Top) / 2
  lblResult.Move 0, txtSQL.Top + txtSQL.Height + 60
  txtResult.Move 0, lblResult.Top + lblResult.Height + 60, ScaleWidth, ScaleHeight - (lblResult.Top + lblResult.Height + 60)
  cmdEnd.Width = ScaleWidth - cmdEnd.Left - 60
End Sub

' End
Private Sub cmdEnd_Click()
  Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Clipboard.Clear
  Clipboard.SetText txtResult
End Sub


