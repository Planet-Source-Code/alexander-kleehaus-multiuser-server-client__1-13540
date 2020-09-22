VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Client"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox txtCommand 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   3015
   End
   Begin VB.TextBox txtIP 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   960
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Remote IP :"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Remote Host has local IP"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Value           =   -1  'True
      Width           =   3015
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   150
      Width           =   1335
   End
   Begin VB.TextBox txtReceived 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox txtMessage 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   3015
   End
   Begin MSWinsockLib.Winsock wsMain 
      Left            =   2760
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   2400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   120
      Top             =   1440
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Zentriert
      Caption         =   "Send Command"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      Caption         =   "Send Message"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3120
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label3 
      Caption         =   "Name:"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   180
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Caption         =   "Received From Server"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   3015
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

Public server_answer As String

Private Sub Command1_Click()
        
  If Command1.Caption = "Connect" Then
    
    If txtUserName.Text = "" Then
      MsgBox "You need to type your username!", vbCritical, "Unable to complete"
      Exit Sub
    End If
    If txtIP.Text = "" Then
      MsgBox "IP-Address not valid!"
      Exit Sub
    End If
    wsMain.RemoteHost = txtIP.Text
    wsMain.Connect
    Do Until wsMain.State = 7
      ' 0 is closed, 9 is error
      If wsMain.State = 0 Or wsMain.State = 9 Then
        MsgBox "Error in connecting!", vbCritical, "Winsock Error"
        ' there was an error, so let's leave
        wsMain.Close
        Exit Sub
      End If
      DoEvents  'don't freeze the system!
    Loop
    ' "log-in":
    wsMain.SendData "U" & Chr(1) & txtUserName.Text
    txtUserName.Enabled = False
    txtMessage.Enabled = True
    Command1.Caption = "Disconnect"
      
  Else
    
    wsMain.Close
    Command1.Caption = "Connect"
    txtReceived = ""
  
  End If
  
End Sub

Private Sub Form_Load()

  Call Option1_Click

End Sub

Private Sub Option1_Click()

  Option1.Value = True
  Option2.Value = False

  txtIP.BackColor = &H8000000B
  txtIP.Text = wsMain.LocalIP
  txtIP.Enabled = False
  
End Sub

Private Sub Option2_Click()

  Option1.Value = False
  Option2.Value = True

  txtIP.BackColor = &H80000005
  txtIP.Text = ""
  txtIP.Enabled = True
  
End Sub

Private Sub txtCommand_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If wsMain.State = sckConnected Then
      wsMain.SendData "r" & Chr(1) & txtCommand.Text
      txtCommand.Text = ""
      KeyAscii = 0
    Else
      MsgBox "Es existiert momentan keine Verbindung!"
      txtMessage.Text = ""
    End If
  End If
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If wsMain.State = sckConnected Then
      wsMain.SendData "t" & Chr(1) & txtMessage.Text
      txtMessage.Text = ""
      KeyAscii = 0
    Else
      MsgBox "Es existiert momentan keine Verbindung!"
      txtMessage.Text = ""
    End If
  End If
End Sub

Private Sub wsMain_Close()
        
  txtReceived.SelStart = Len(txtReceived.Text)
  txtReceived.SelText = "Connection to Server lost" & vbCrLf
  
End Sub

Private Sub wsMain_DataArrival(ByVal bytesTotal As Long)
  
  Dim Data As String, CtrlChar As String
  wsMain.GetData Data
  CtrlChar = Left(Data, 1) ' Let's get the first char
  Data = Mid(Data, 3)      ' Then cut it off
  Select Case LCase(CtrlChar)   ' Check what it is
    Case "m"   ' Do stuff depending on it
      MsgBox Data, vbInformation, "Msg from server"
    Case "c"
      Me.Caption = "Client - " & Data
    Case "r"
      server_answer = Data
      txtReceived.SelStart = Len(txtReceived.Text)
      txtReceived.SelText = Data & vbCrLf
    Case Else
      txtReceived.SelStart = Len(txtReceived.Text)
      txtReceived.SelText = Data & vbCrLf
  End Select
  
End Sub

Private Sub wsMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  MsgBox "Winsock Error: " & Number & vbCrLf & Description, vbCritical, "Winsock Error"
End Sub

Private Function Senden(tex As String) As String

  Dim antwort As String
  Dim tax As String
  On Error Resume Next

  tax = tex
  If tax = "" Then Exit Function
  server_answer = ""
  wsMain.SendData "r" & Chr(1) & tex
  Timer1.Enabled = True
  
  DoEvents
  Do Until server_answer <> ""
    DoEvents
  Loop
  
  If server_answer = "Time-Out" Then
    MsgBox "Server Time Out!"
  End If
  
  Senden = server_answer
  
  Exit Function
  
errhand:

  tex = Err.Description
  'Errn = Err
  'If Errn = 10048 Then
  '  Resume
  'Else
    MsgBox tex, 16, "Fehler im WinSock"
    Resume
  'End If

End Function

Private Sub Timer1_Timer()

  server_answer = "Server-Time out"
  Timer1.Enabled = False

End Sub


