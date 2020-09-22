VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Server"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.ListView lvwUsers 
      Height          =   1935
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdMsgBox 
      Caption         =   "Popup Message Box"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   5280
      Width           =   2535
   End
   Begin VB.CommandButton cmdCaption 
      Caption         =   "Set their Caption"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5280
      Width           =   2535
   End
   Begin VB.TextBox txtReceived 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3840
      Width           =   5175
   End
   Begin VB.TextBox txtSendMessage 
      Height          =   285
      Left            =   3120
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txtErrors 
      Height          =   1335
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2040
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock wsArray 
      Index           =   0
      Left            =   4920
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   2500
   End
   Begin MSWinsockLib.Winsock wsListen 
      Left            =   120
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   2400
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5280
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   255
      Left            =   1800
      TabIndex        =   15
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblLocalPort 
      Caption         =   "Local Port:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblLocalIP 
      Caption         =   "Local IP:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblServerName 
      Caption         =   "Server Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Zentriert
      Caption         =   "(shift-enter to broadcast)"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Zentriert
      Caption         =   "Received"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   5175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Zentriert
      Caption         =   "Send Message"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      Caption         =   "Error Log"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Caption         =   "Users"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
' Commands:
'------------------------------------------------------------------------------
' NewBuffer
' - Creates a new Buffer for Selections
'
' DeleteBuffer|[Key]
' - Deletes a specified Buffer
'
' GetBuffer
' - Shows the Keys of all existing Buffers
'
' Selection|[Key]|[String]
' - Selects all words in a specified Buffer with wildcard (*) search
'
' GetItem|[Key]|[Index]
' - Gets the value of the specified Buffer and its inde
'
' Sample:   NewBuffer
'           Selection|1|A*
'           GetItem|1|1
'           NewBuffer
'           GetBuffer
'           DeleteBuffer|1
'           DeleteBuffer|2
'------------------------------------------------------------------------------

Private Sub cmdCaption_Click()

  Dim User As Integer
  ' Get Username to send to
  User = RetrieveUser(lvwUsers.SelectedItem.Text)
  If User = -1 Then
    MsgBox "Invalid User!", vbCritical, "Error"
    Exit Sub
  End If
  wsArray(User).SendData "c" & Chr(1) & InputBox("What do you want to have their caption set to?", "Alter Caption", "Hi!")
  
End Sub

Private Sub cmdMsgBox_Click()

  Dim User As Integer
  ' Get Username to send to
  User = RetrieveUser(lvwUsers.SelectedItem.Text)
  If User = -1 Then
    MsgBox "Invalid User!", vbCritical, "Error"
    Exit Sub
  End If
  wsArray(RetrieveUser(lvwUsers.SelectedItem.Text)).SendData "m" & Chr(1) & InputBox("What do you want to have displayed on their machine?", "Popup MsgBox", "Hi!")

End Sub

Private Sub Form_Load()

  Dim colHead As ColumnHeader
  
  Set colHead = lvwUsers.ColumnHeaders.Add(, , "Name", lvwUsers.Width / 3)
  Set colHead = lvwUsers.ColumnHeaders.Add(, , "IP", lvwUsers.Width / 3)
  Set colHead = lvwUsers.ColumnHeaders.Add(, , "Port", lvwUsers.Width / 3)

  Label6.Caption = wsListen.LocalHostName
  Label7.Caption = wsListen.LocalIP
  Label8.Caption = wsListen.LocalPort
  
  FillDB
  
  wsListen.Listen  ' make it listen
  
  Set DB = New ServerDatabase
  
End Sub

Private Sub txtSendMessage_KeyDown(KeyCode As Integer, Shift As Integer)

  Dim Index As Integer
  Dim x As Integer
  
  'First, check to make sure someone's logged in
  If lvwUsers.ListItems.count = 0 And KeyCode = 13 Then
  
    'Display popup
    MsgBox "Nobody to send to!", vbExclamation, "Cannot send"
    
    'Clear input
    txtSendMessage.Text = ""
    Exit Sub
  End If

  ' If it was enter and shift wasn't pressed, then...
  If KeyCode = 13 And Shift = 0 Then
    ' Get Username to send to
    Index = RetrieveUser(lvwUsers.SelectedItem.Text)
    ' RetrieveUser returns -1 if the user wasn't found
    If Index = -1 Then
      Exit Sub
    End If
    ' format the message
    wsArray(Index).SendData "t" & Chr(1) & txtSendMessage.Text
    ' Blank the input
    txtSendMessage.Text = ""
  
  ElseIf KeyCode = 13 And Shift = 1 Then
      
    ' Loop through the users.
    ' There's better ways of doing this
    Dim User As User
    For Each User In DB.Users
      'Send the message
      wsArray(User.SocketIndex).SendData "t" & Chr(1) & txtSendMessage.Text
    
      ' Don't know why this needs to be
      ' in here to work - someone tell me?
      'DoEvents
    Next
    txtSendMessage.Text = ""
  End If

End Sub

Private Function RetrieveUser(UserName As String) As Integer
  
  Dim x As Integer

  'Check to see if nothing was selected
  If UserName = "" Then
      
    'OK, nothing selected, let's see how full
    ' the list is!
    If lvwUsers.SelectedItem.Index = 0 Then
        
      'Nothing in the list, so return -1
      RetrieveUser = -1
      Exit Function
    End If
    
    'If there is something in the list, send it to
    ' the first one =)
    UserName = lvwUsers.ListItems(1)
  End If
  
  ' Count through the users
  For x = 1 To DB.Users.count
      
    'Check username to see if it is the right one
    If DB.Users(x).Name = UserName Then
    
      'Ok, this is our man, so let's return his
      ' winsock index
      RetrieveUser = DB.Users(x).SocketIndex
      Exit Function
    End If
  Next x
  RetrieveUser = -1
  
End Function

Private Sub txtSendMessage_KeyPress(KeyAscii As Integer)
  'Let's get rid of the annoying beep =)
  If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub wsArray_Close(Index As Integer)
   
  Dim x As Integer
  ' Let's cycle through the list, looking for their
  ' name
  
  With DB.Users
    For x = 1 To .count
      If .Item(x).SocketIndex = Index Then
        .Remove x
        lvwUsers.ListItems.Remove x
        wsArray(Index).Close
        Exit Sub
      End If
    Next
  End With
  
End Sub

Private Sub wsArray_DataArrival(Index As Integer, ByVal bytesTotal As Long)

  Dim Data As String, CtrlChar As String
  Dim lvwItem As ListItem
  Dim Name As String
  Dim IP As String
  Dim Port As Integer
  
  wsArray(Index).GetData Data
  
  ' Our format for our messages is this:
  ' CtrlChar & chr(1) & <info>
  If InStr(1, Data, Chr(1)) <> 2 Then
  ' If the 2nd char isn't chr(1), we know we have a prob
    MsgBox "Unknown Data Format: " & vbCrLf & Data, vbCritical, "Error receiving"
    ' Make sure to leave the sub so it doesn't
    ' try to process the invalid info!
    Exit Sub
  End If
  
  'Retrieve First Character
  CtrlChar = Left(Data, 1)

  'Make sure to trim it, and chr(1), off
  Data = Mid(Data, 3)
  
  ' Check what it is, without regard to case
  Select Case LCase(CtrlChar)
      
    'This is to display a msgbox.
    ' I didn't enable the ability on the clients --
    '  for obvious reasons ;)
    Case "m"
      MsgBox Data, vbInformation, "Msg from client"
    
    'This is to change the caption.
    ' I didn't enable the ability on the clients --
    '  for obvious reasons ;)
    Case "c"
      Me.Caption = "Server - " & Data
    
    'This is their "login" key
    Case "u"
    
      'Add their name to the array
      Name = Data
      IP = wsArray(Index).RemoteHostIP
      Port = wsArray(Index).RemotePort
      DB.Users.Add Name, IP, Port, Index
  
      'Add their name to the list
      Set lvwItem = lvwUsers.ListItems.Add
      lvwItem.Text = Name
      lvwItem.SubItems(1) = IP
      lvwItem.SubItems(2) = Port
      
      ' We need to remember that both
      ' the winsock index and the user array
      ' index correspond.  So you can find a
      ' users name by going "Users(<winsock index>)"
      ' or you can find the winsock index with
      ' a text name by cycling through the array.
      ' That's what the function "RetrieveUser"
      ' does - gets their winsock index from their
      ' username
        
    ' If all else fails, print it to output =)
    Case "r"
      wsArray(Index).SendData "r" & Chr(1) & ClientCall(Data, Index + 1)
    Case Else
      Dim User As User
      For Each User In DB.Users
        If User.SocketIndex = Index Then
          Name = User.Name
          Exit For
        End If
      Next
      txtReceived.SelStart = Len(txtReceived.Text)
      txtReceived.SelText = Name & "> " & Data & vbCrLf
  End Select
  
End Sub

Private Sub wsArray_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
  ' This sets the "cursor" to the end of the textbox
  txtErrors.SelStart = Len(txtErrors.Text)
  
  ' This inserts the error message at the "cursor"
  txtErrors.SelText = "wsArray(" & Index & ") - " & Number & " - " & Description & vbCrLf
  
  ' Close it =)
  wsArray(Index).Close
  
End Sub

Private Sub wsListen_ConnectionRequest(ByVal requestID As Long)
  
  Dim Index As Integer
  Index = FindOpenWinsock
  
  ' Accept the request using the created winsock
  wsArray(Index).Accept requestID
End Sub

Private Sub wsListen_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
  ' This sets the "cursor" to the end of the textbox
  txtErrors.SelStart = Len(txtErrors.Text)
  
  ' This inserts the error message at the "cursor"
  txtErrors.SelText = "wsListen - " & Number & " - " & Description & vbCrLf
End Sub

Private Function FindOpenWinsock()

  Static LocalPorts As Integer  ' Static keeps the
                                ' variable's state
  Dim x As Integer
  
  For x = 0 To wsArray.UBound
    If wsArray(x).State = 0 Then
        
      ' We found one that's state is 0, which
      '  means "closed", so let's use it
      FindOpenWinsock = x
      
      ' make sure to leave function
      Exit Function
    End If
  Next x

  '  OK, none are open so let's make one
  Load wsArray(wsArray.UBound + 1)
  
  '  Let's make sure we don't get conflicting local ports
  LocalPorts = LocalPorts + 1
  wsArray(wsArray.UBound).LocalPort = wsArray(wsArray.UBound).LocalPort + LocalPorts
  
  '  and then let's return it's index value
  FindOpenWinsock = wsArray.UBound

End Function
