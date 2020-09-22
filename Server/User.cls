VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Buffers"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mName As String           'lokale Kopie
Private mIP As String             'lokale Kopie
Private mPort As Integer          'lokale Kopie
Private mSocketIndex As Integer   'lokale Kopie
Private mBuffers As Buffers

Public Property Get Buffers() As Buffers
  If mBuffers Is Nothing Then
    Set mBuffers = New Buffers
  End If
  Set Buffers = mBuffers
End Property

Public Property Set Buffers(vData As Buffers)
  Set mBuffers = vData
End Property
Private Sub Class_Terminate()
  Set mBuffers = Nothing
End Sub


Public Property Let Name(ByVal vData As String)
  'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
  'Syntax: X.IP = 5
  mName = vData
End Property

Public Property Get Name() As String
  'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
  'Syntax: Debug.Print X.IP
  Name = mName
End Property

Public Property Let Port(ByVal vData As Integer)
  'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
  'Syntax: X.Port = 5
  mPort = vData
End Property

Public Property Get Port() As Integer
  'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
  'Syntax: Debug.Print X.Port
  Port = mPort
End Property

Public Property Let IP(ByVal vData As String)
  'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
  'Syntax: X.IP = 5
    mIP = vData
End Property

Public Property Get IP() As String
  'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
  'Syntax: Debug.Print X.IP
  IP = mIP
End Property

Public Property Let SocketIndex(ByVal vData As String)
  'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
  'Syntax: X.IP = 5
    mSocketIndex = vData
End Property

Public Property Get SocketIndex() As String
  'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
  'Syntax: Debug.Print X.IP
  SocketIndex = mSocketIndex
End Property

