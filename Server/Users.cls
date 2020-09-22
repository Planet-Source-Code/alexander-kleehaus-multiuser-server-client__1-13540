VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"User"
Attribute VB_Ext_KEY = "Member0" ,"User"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Lokale Variable für eine Auflistung.
Private colUsers As Collection

Public Function Add(Name As String, IP As String, Port As Integer, SocketIndex As Integer, Optional sKey As String) As User
    
  'neues Objekt erstellen
  Dim newUser As User
  Set newUser = New User

  'die Eigenschaften festlegen, die der Methode übergeben werden
  newUser.Name = Name
  newUser.IP = IP
  newUser.Port = Port
  newUser.SocketIndex = SocketIndex
  
  Dim newBuffers As Buffers
  Set newUser.Buffers = New Buffers
  
  If Len(sKey) = 0 Then
    colUsers.Add newUser
  Else
    colUsers.Add newUser, sKey
  End If

  'das erstellte Objekt zurückgeben
  Set Add = newUser
  Set newUser = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As User
Attribute Item.VB_UserMemId = 0
  'Wird verwendet, wenn auf ein Element der Auflistung zugegriffen
  'wird. vntIndexKey enthält entweder den Index oder den Schlüssel.
  'Daher wird der Parameter als Variant deklariert.
  'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = colUsers(vntIndexKey)
End Property

Public Property Get Count() As Long
  'Wird verwendet, wenn die Anzahl der Elemente
  'in der Auflistung ermittelt werden soll.
  'Syntax: Debug.Print x.Count
  Count = colUsers.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
  'Wird verwendet, wenn ein Element aus der Auflistung
  'entfernt werden soll. vntIndexKey enthält entweder
  'den Index oder den Schlüssel. Daher wird der Parameter
  'als Variant deklariert.
  'Syntax: x.Remove(xyz)
  colUsers.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  'Diese Eigenschaft ermöglicht, die Elemente der
  'Auflistung mit For...Each zu durchlaufen.
  Set NewEnum = colUsers.[_NewEnum]
End Property

Private Sub Class_Initialize()
  'Erstellt die Auflistung, wenn die Klasse erstellt wird.
  Set colUsers = New Collection
End Sub

Private Sub Class_Terminate()
  'Entfernt die Auflistung aus dem Speicher, wenn die Klasse beendet wird.
  Set colUsers = Nothing
End Sub

