Attribute VB_Name = "Module1"
Option Explicit

Public DB As ServerDatabase

Public Function ClientCall(Data As String, Index As Integer) As String

  Dim Methode As String
  Dim mArray() As String
  Dim ret As String
  
  mArray = Split(Data, "|")
  Methode = mArray(0)
  
  Select Case Methode
    Case "GetBuffer"
      ret = GetBuffer(Index)
    Case "NewBuffer"
      ret = NewBuffer(Index)
    Case "DeleteBuffer"
      If TestArguments(mArray, 1, ret) = True Then ret = DeleteBuffer(Index, mArray(1))
    Case "Selection"
      If TestArguments(mArray, 2, ret) = True Then ret = Selection(Index, mArray(1), mArray(2))
    Case "GetItem"
      If TestArguments(mArray, 2, ret) = True Then ret = GetItem(Index, mArray(1), mArray(2))
    Case Else
      ret = "Unknown Command"
  End Select

  ClientCall = ret

End Function

'******************************************************************************
'* Functions
'******************************************************************************

Private Function NewBuffer(ByVal Index As Integer) As String
  
  Dim Key As String
  
  With DB.Users(Index)
    Key = GetKeyBuffer(Index)
    .Buffers.Add "B" & Key
    .Buffers("B" & Key).BufferKey = Key
  End With
  
  NewBuffer = "Buffer " & Key & " created"
    
End Function

Private Function GetKeyBuffer(ByVal Index As Integer) As Integer

  Dim mBuffer As Buffer
  Dim tmp As Integer
  Dim nr As Integer
  
  With DB.Users(Index)
    If .Buffers.count = 0 Then
      nr = 1
    Else
      For Each mBuffer In .Buffers
        tmp = mBuffer.BufferKey
        If tmp > nr Then
          nr = tmp
        End If
      Next
      nr = nr + 1
    End If
  End With

  GetKeyBuffer = nr
  
End Function

Private Function DeleteBuffer(ByVal Index As Integer, ByVal Key As String)

  Dim mBuffer As Buffer
  Dim tmp As String
  Dim deleted As Boolean
  
  If TestBuffer(Index, BufIndex, tmp) = True Then
    With DB.Users(Index)
      For Each mBuffer In .Buffers
        If mBuffer.BufferKey = Key Then
          .Buffers.Remove "B" & Key
          deleted = True
          tmp = "Buffer " & Key & " deleted"
        End If
      Next
      If deleted = False Then
        tmp = "Can't allocate Buffer " & Key
      End If
    End With
  End If
  
  DeleteBuffer = tmp

End Function

Private Function GetBuffer(ByVal Index As Integer) As String

  Dim mBuffer As Buffer
  Dim tmp As String
  
  If TestBuffer(Index, BufIndex, tmp) = True Then
    With DB.Users(Index)
      tmp = "Buffer "
      For Each mBuffer In .Buffers
        tmp = tmp & mBuffer.BufferKey & ", "
      Next
      tmp = Mid(tmp, 1, Len(tmp) - 2)
    End With
  End If
  
  GetBuffer = tmp
  
End Function

Private Function Selection(ByVal Index As Integer, ByVal BufIndex As Integer, ByVal criteria As String) As String

  Dim mBuffer As Buffer
  Dim tmp As String
  Dim i As Long
  Dim length As Integer
  Dim wildcardfront As Boolean
  Dim wildcardback As Boolean
  
  If TestBuffer(Index, BufIndex, tmp) = True Then
    length = Len(criteria)
    
    If Left$(criteria, 1) = "*" Then
      wildcardfront = True
      length = length - 1
      criteria = Mid$(criteria, 2)
    End If
    If Right$(criteria, 1) = "*" Then
      wildcardback = True
      length = length - 1
      criteria = Mid$(criteria, 1, length)
    End If
    
    With DB.Users(Index).Buffers(BufIndex)
      .Clear
      '------------------------------------------------------------------------
      ' WildCards
      If wildcardfront = False And wildcardback = False Then
        For i = 0 To UBound(DBArray)
          tmp = DBArray(i)
          If Len(tmp) = length Then
            If tmp = criteria Then
              .Add tmp
            End If
          End If
        Next i
      ElseIf wildcardfront = False And wildcardback = True Then
        For i = 0 To UBound(DBArray)
          tmp = DBArray(i)
          If Len(tmp) >= length Then
            If Left$(tmp, length) = criteria Then
              .Add tmp
            End If
          End If
        Next i
      ElseIf wildcardfront = True And wildcardback = False Then
        For i = 0 To UBound(DBArray)
          tmp = DBArray(i)
          If Len(tmp) >= length Then
            If Right$(tmp, length) = criteria Then
              .Add tmp
            End If
          End If
        Next i
      Else
        For i = 0 To UBound(DBArray)
          tmp = DBArray(i)
          If Len(tmp) >= length Then
            If InStr(tmp, criteria) > 0 Then
              .Add tmp
            End If
          End If
        Next i
      End If
      '------------------------------------------------------------------------
      tmp = .count & " Items selected"
    End With
  End If
  
  Selection = tmp
  
End Function

Private Function GetItem(ByVal Index As Integer, ByVal BufIndex As Integer, ByVal ItemIndex As Long)

  Dim tmp As String
  
  'Gets value from Buffer
  If TestBuffer(Index, BufIndex, tmp) = True Then
    tmp = DB.Users(Index).Buffers(BufIndex).Item(ItemIndex)
  End If
  
  GetItem = tmp

End Function

'******************************************************************************
'* Test-Functions
'******************************************************************************

Private Function TestArguments(ByRef arr() As String, ByVal count As Integer, ByRef error As String) As Boolean

  Dim value As Integer
  
  value = UBound(arr)
  
  'Test if there are all needed arguments
  If value < count Then
    error = "Missing argument"
  ElseIf value > count Then
    error = "Too many arguments"
  Else
    TestArguments = True
  End If

End Function

Private Function TestBuffer(ByVal Index As Integer, ByVal BufIndex As Integer, ByRef error As String) As Boolean
  
  Dim mBuffer As Buffer
  
  TestBuffer = False
  
  'Test if there are any Buffers allocated
  If DB.Users(Index).Buffers.count = 0 Then
    error = "No Buffers allocated"
    Exit Function
  End If
  
  'Test if the BufIndex (Key) exists
  For Each mBuffer In DB.Users(Index).Buffers
    If mBuffer.BufferKey = BufIndex Then
      TestBuffer = True
      Exit Function
    End If
  Next
      
  error = "Buffer " & BufIndex & " is not valid"
    
End Function
