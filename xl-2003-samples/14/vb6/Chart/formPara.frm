VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formPara 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Parameters of Chart"
   ClientHeight    =   2295
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   4275
   Icon            =   "formPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   2295
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1740
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Redraw Chart"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   1740
      Width           =   1815
   End
   Begin MSComctlLib.Slider SliderY 
      Height          =   630
      Left            =   2280
      TabIndex        =   7
      Top             =   420
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1111
      _Version        =   393216
      Min             =   10
      Max             =   50
      SelStart        =   10
      Value           =   10
   End
   Begin MSComctlLib.Slider SliderX 
      Height          =   630
      Left            =   240
      TabIndex        =   0
      Top             =   420
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1111
      _Version        =   393216
      Min             =   10
      Max             =   50
      SelStart        =   10
      Value           =   10
   End
   Begin VB.Label LabelY 
      Caption         =   "Y Frequency"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   180
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "1"
      Height          =   195
      Left            =   2400
      TabIndex        =   5
      Top             =   1020
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "5"
      Height          =   195
      Left            =   3720
      TabIndex        =   4
      Top             =   1020
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "5"
      Height          =   195
      Left            =   1680
      TabIndex        =   3
      Top             =   1020
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "1"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   1020
      Width           =   255
   End
   Begin VB.Label LabelX 
      Caption         =   "X Frequency"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   180
      Width           =   1575
   End
End
Attribute VB_Name = "formPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Tag = "ok"
  Hide
End Sub

Private Sub Command2_Click()
  Tag = "cancel"
  Hide
End Sub

Private Sub Sliderx_Change()
  LabelX = "X frequency: " & SliderX / 10
End Sub
Private Sub Sliderx_Scroll()
  LabelX = "Y frequency: " & SliderX / 10
End Sub
Private Sub Slidery_Change()
  LabelY = "Y frequency: " & SliderY / 10
End Sub
Private Sub Slidery_Scroll()
  LabelY = "Y frequency: " & SliderY / 10
End Sub
