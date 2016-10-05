VERSION 5.00
Begin VB.Form formOLE 
   ClientHeight    =   4065
   ClientLeft      =   2295
   ClientTop       =   3030
   ClientWidth     =   5355
   ControlBox      =   0   'False
   Icon            =   "formOLE.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   4065
   ScaleWidth      =   5355
   WindowState     =   2  'Maximiert
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   -2000
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.OLE OLE1 
      Class           =   "Excel.Sheet.8"
      Height          =   3375
      Left            =   0
      OLETypeAllowed  =   1  'Eingebettet
      SizeMode        =   1  'Strecken
      SourceDoc       =   "H:\Code\vb-5\ActiveX-Automation\Excel\ExcelGrafik.xls"
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
   Begin VB.Menu menuMain 
      Caption         =   "3D Chart"
      Index           =   1
      NegotiatePosition=   1  'Links
      Begin VB.Menu menuPara 
         Caption         =   "Change &Chart parameters"
      End
      Begin VB.Menu menuSettings 
         Caption         =   "Change &Excel settings"
      End
      Begin VB.Menu menuClose 
         Caption         =   "&Deactivate Excel "
      End
      Begin VB.Menu menuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu menuEnd 
         Caption         =   "&Quit"
      End
   End
End
Attribute VB_Name = "formOLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'VB6\Chart\formOLE.frm
Option Explicit


Dim wb As Workbook

' initialize: load Excel file, save reference to workbook
' in the variable wb, insert data, show result in
' OLE control
Private Sub Form_Load()
  Dim xl As Object, win As Window
  On Error Resume Next
  ' ChDrive App.Path
  ' ChDir App.Path
  Me.OLE1.Visible = False  'hidden
  formWait.Show            'please wait ...
  MousePointer = vbHourglass
  With Me
    .OLE1.CreateEmbed App.Path + "\ActiveX_Chart.xls"
    Set xl = .OLE1.object.Application
    ' loop through all Excel windows,
    ' search for new window
    For Each win In xl.Windows
      If win.Parent.Title = "ActiveX_Chart_keyword" Then
        ' got it!
        Set wb = win.Parent
        Exit For
      End If
    Next
  End With
  
  ' if an error has occured, is Excel
  ' probably not available
  If Err <> 0 Then
    MsgBox "An error has occurred. " _
      & "This program is stopped. You need Excel 2000 " _
      & "to test this sample."
    Unload Me
  End If
  PlotChart
  Me.OLE1.Visible = True
  MousePointer = 0
  formWait.Hide
End Sub


' change graphics parameter
Private Sub menuPara_Click()
  Dim xfreq, yfreq
  ' show form to change parameters
  With formPara
    xfreq = .SliderX: yfreq = .SliderY
    .Show vbModal
    If .Tag = "cancel" Then
      .SliderX = xfreq: .SliderY = yfreq
      Exit Sub
    End If
  End With
  PlotChart  'redraw chart
End Sub

' change Excel settings
Private Sub menuSettings_Click()
  Me.OLE1.DoVerb
End Sub

' print chart
Private Sub menuPrint_Click()
  On Error Resume Next
  wb.Sheets("chart").PrintOut
  If Err <> 0 Then
    MsgBox "Beim Versuch, das Diagramm zu drucken, ist ein Fehler aufgetreten"
  End If
  ' alternative: call procedure PrintChart
  '  in Module1 of the workbook
  ' Dim pname$
  '  pname = wb.Name & "!Module1.PrintChart"
  '  wb.Application.Run pname
End Sub


' create test data and transfer via clipboard
' into the sheet (this is the fastest way to do it)
Sub PlotChart()
  Dim xfreq, yfreq
  Dim x#, y#, z#, data$
  xfreq = formPara.SliderX
  yfreq = formPara.SliderY
  ' calculate new data
  For y = 0 To 2.00001 Step 0.1
    For x = 0 To 2.00001 Step 0.1
      z = Sin(x * xfreq / 10) + Sin(y * yfreq / 10)
      data = data & DecimalPoint(Str(z)) & vbTab
    Next x
    data = data & vbCr
  Next y
  Clipboard.Clear
  Clipboard.SetText data
  wb.Sheets("table").Paste wb.Sheets("table").Cells(2, 2)
  ' show the chart, not the worksheet
  wb.Sheets("chart").Activate
  ' Activate does not suffice (for whatever reason ...)
  wb.Sheets("table").Visible = False
End Sub

' replace comma by point (only neccessary for some regional settings)
Private Function DecimalPoint$(x$)
  DecimalPoint = Replace(x, ",", ".")
End Function

' deactive Excel
Private Sub menuClose_Click()
  Text1.SetFocus
End Sub


' end
Private Sub menuEnd_Click()
  Unload Me
  End
End Sub

' end
Private Sub Form_Unload(Cancel As Integer)
  Set wb = Nothing  'Excel is no longer needed
  End
End Sub

' resize OLE control
Private Sub Form_Resize()
  If WindowState = vbMinimized Then Exit Sub
  OLE1.Width = ScaleWidth
  OLE1.Height = ScaleHeight
End Sub

