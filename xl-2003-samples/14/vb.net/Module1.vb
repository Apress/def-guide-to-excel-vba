' to use late binding
Option Strict Off
Module Module1
  Sub Main()
    ' process Excel file
    process_xl_file()
    ' close Excel
    GC.Collect()
    GC.WaitForPendingFinalizers()
    GC.Collect()
    GC.WaitForPendingFinalizers()
    ' end this program
    Console.WriteLine("Hit Return to end")
    Console.ReadLine()
  End Sub

  Sub process_xl_file()
    Dim i, j As Integer
    Dim xl, wb, ws As Object
    Dim fname As String
    fname = IO.Path.Combine(Environment.CurrentDirectory, "..\sample.xls")
    wb = GetObject(fname)
    xl = wb.Application
    ' xl.Visible = True 'if you want to see Excel
    ' wb.NewWindow()
    ws = wb.Sheets(1)
    For i = 1 To 3
      For j = 1 To 3
        Console.WriteLine("Cell in line {0} / column {1} ={2}", i, j, ws.Cells(i, j).Value)
      Next
    Next
    ws.Cells(4, 1).Value = Now
    ' wb.Windows(wb.Windows.Count).Close()
    wb.Save()
    wb.Close()
    If xl.Workbooks.Count = 0 Then xl.Quit()
  End Sub
End Module
