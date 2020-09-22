Attribute VB_Name = "Module2"
Option Explicit

Public DBArray() As String

Public Sub FillDB()

  Dim handl As Integer
  Dim tmp As String
  Dim cnt As Long
  
  ReDim DBArray(26870)
  
  cnt = 0
  Open App.Path & "\Words.txt" For Input As #1
    Do While Not EOF(1)
      Line Input #1, DBArray(cnt)
      cnt = cnt + 1
    Loop
  Close #1

End Sub
