VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Buffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mBufferKey      As Integer    'lokale Kopie

Private m_Count         As Long       'this will contain the number of items
Private m_Array()       As String     'Main Array

Public Function Clear()
  
  ReDim m_Array(0) As String
  m_Count = 0

End Function

Public Function Count() As Long
  
  m_Count = UBound(m_Array)
  Count = m_Count

End Function

Public Function Add(ByVal Item As String)
  
  m_Count = m_Count + 1
  ReDim Preserve m_Array(m_Count) As String
  m_Array(m_Count) = Item

End Function

Private Function Remove(ByVal m_Index As Long)

  Dim Counter As Integer

  If (m_Index >= 0) And (m_Index <= m_Count) Then
    If m_Count = 0 Then
      m_Array(0) = 0
      Exit Function
    End If
      
    m_Count = m_Count - 1
      
    If m_Index < m_Count Then
      For Counter = m_Index To m_Count
        m_Array(Counter) = m_Array(Counter + 1)
        m_Array(Counter) = m_Array(Counter + 1)
      Next Counter
    End If
    
    ReDim Preserve m_Array(m_Count) As String
  End If

End Function

Public Function Item(ByVal m_Index As Long) As String
Attribute Item.VB_UserMemId = 0

  If (m_Index >= 0) And (m_Index <= m_Count) Then
    Item = m_Array(m_Index)
  Else
    Item = "Index not valid"
  End If
  
End Function

Private Sub Class_Initialize()
  
  Clear

End Sub

Private Sub Class_Terminate()

  Erase m_Array

End Sub

Public Property Let BufferKey(ByVal vData As Integer)
  'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
  'Syntax: X.IP = 5
  mBufferKey = vData
End Property

Public Property Get BufferKey() As Integer
  'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
  'Syntax: Debug.Print X.IP
  BufferKey = mBufferKey
End Property


