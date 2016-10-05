VERSION 5.00
Begin VB.Form formWait 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Please wait..."
   ClientHeight    =   2790
   ClientLeft      =   1215
   ClientTop       =   2070
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "formWait.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   2790
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      Height          =   2655
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "formWait.frx":0442
      Top             =   120
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "formWait.frx":0575
      Top             =   180
      Width           =   480
   End
End
Attribute VB_Name = "formWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

