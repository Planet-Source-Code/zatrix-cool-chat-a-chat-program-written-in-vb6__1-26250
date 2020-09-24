VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCheckEmail 
   BackColor       =   &H00FF8080&
   Caption         =   "Cool Chat - Check Email"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4380
   Icon            =   "CheckEmail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   4380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Port Connection Status:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   3975
      Begin VB.ListBox lstStatus 
         Height          =   645
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Port Status Messages"
         Top             =   240
         Width           =   3735
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Status Progress Bar"
         Top             =   960
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Text            =   "110"
      ToolTipText     =   "POP3 Port - Default is 110"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Pop3 Username (Usually same as ISP username)"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Usually: mail.yourisp.com , pop.yourisp.com , pop3.yourisp.com (Not for AOL users)"
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CheckBox chkLogStatus 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Log Port Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      TabIndex        =   4
      ToolTipText     =   "Check this if you want to see the responses from the POP3 server"
      Top             =   2280
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3600
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Pop3 Password (Usually same as ISP password)"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H00FF8080&
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H00FF8080&
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H00FF8080&
      Caption         =   "POP3 Server:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H00FF8080&
      Caption         =   "POP3 Port:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.Menu mnuCheckEmail 
      Caption         =   "&Check Email"
   End
End
Attribute VB_Name = "frmCheckEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Constant used to limit the number of list
'box items before removing the oldest
Private Const MAX_LIST_ITEMS = 100

'Variables to store waiting message information
Private mnMessageCount As Integer
Private mlMessageChars As Long

'This variable hold the most recent command
'that was sent to the POP3 Server
Private msCommand As String

'State of the Log Port Status Check Box
Private mbLogStatus As Boolean

'Variable to track state of socket data reception
Private mbGotData As Boolean

'Array for socket state descriptions
Private mvntSocketState As Variant


Private Sub chkLogStatus_Click()

'Update the state variable based on the
'value of the check box
If chkLogStatus.Value = vbChecked Then
    mbLogStatus = True
Else
    mbLogStatus = False
End If
End Sub

Private Sub mnuCheckEmail_Click()

Dim nMonitorPort As Integer
Dim sServer As String

    'Get the Value used for the server name
    sServer = txtServer.Text
    If Len(sServer) = 0 Then
        MsgBox "Plese enter the POP3 server name"
        Exit Sub
    End If
    
'Get the value used for the port connection
nMonitorPort = CInt(txtPort.Text)
If nMonitorPort = 0 Then
    MsgBox "Please enter the POP3 server and Port (Default = 110)"
    Exit Sub
End If

'Error checking for username and password has been
'left out intentionally to show the error messages
'returned by the POP3 Server

'Disable this command button as the server
'should be treated as a single-threaded for
'a specific user
mnuCheckEmail.Enabled = False

'Connect and get the message count
Call POP3CheckMail

'Enable this command button
mnuCheckEmail.Enabled = True

    
End Sub

Private Sub Form_Load()
'Load the variant arrary with socket states
mvntSocketState = Array("Closed", "Opening", "Listening", "Connection Pending", "Resolving Host", "Host Resolved", "Connecting", "Connected", "Closing", "Error")
                        
'Initialize the log status state to false
mbLogStatus = False


End Sub

Private Sub ShowSocketState()

Dim sTempStr As String
Dim nListCount As Integer
Dim nLoopCtr As Integer

'Check state of port status logging
'If disabled, then exit from this sub
If mbLogStatus = False Then Exit Sub

'Build a string containing the current socket status
sTempStr = "Socket State: " & vbTab & mvntSocketState(Winsock1.State)

'Add the string to the list box
lstStatus.AddItem sTempStr

'Get the index position where the item was added
nListCount = lstStatus.NewIndex

If nListCount > MAX_LIST_ITEMS Then
    'Clean out old list entries
    For nLoopCtr = nListCount - MAX_LIST_ITEMS To 0 Step -1
        lstStatus.RemoveItem nLoopCtr
        Next nLoopCtr
        nListCount = lstStatus.ListCount - 1
        End If
        
'Position at the last item on the list
lstStatus.ListIndex = nListCount

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Winsock1_Close()

'Update the connection status in the list box
Call ShowSocketState

End Sub

Private Sub Winsock1_Connect()

'Update the connection status in the list box
Call ShowSocketState

End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)

'This event should never occur because this program
'does not listen on any port

'Show the connection Request ID
If mbLogStatus = True Then
    lstStatus.AddItem "Unexpected Connection Request"
End If

' The connection is refused by not accepting it

'Update the connection status in the list box
Call ShowSocketState

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

Dim sData As String

'Get the inbound data from the socket
Winsock1.GetData sData, vbString, bytesTotal

'Show the data and its length in the list box
If mbLogStatus = True Then
    lstStatus.AddItem "Data Value:  " & vbTab & sData
    lstStatus.AddItem "Data Length:  " & vbTab & bytesTotal
End If

'Move the data into the form scope command variable
msCommand = Left$(sData, bytesTotal - 2)

'Let other code know we got data
mbGotData = True

'Update the connection status in the list box
Call ShowSocketState

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
'Show the error details
lstStatus.AddItem "Error Number: " & vbTab & Number
lstStatus.AddItem " Error Text: " & vbTab & Description

'Update the connection status in the list box
Call ShowSocketState
    
End Sub

Private Sub Winsock1_SendComplete()

'Show that the send was completed
If mbLogStatus = True Then
    lstStatus.AddItem "Send Complete"
End If

'Update the connection status in the list box
Call ShowSocketState

End Sub


Private Sub POP3CheckMail()

Dim sRemoteAddress
Dim nRemotePort
Dim sUserName
Dim sPassword
Dim bResult

'Get the server address and port
sRemoteAddress = txtServer.Text
nRemotePort = CInt(txtPort.Text)

'Get the username and password
sUserName = txtUser.Text
sPassword = txtPassword.Text

'Set the message count variable to indicate
'a failure to get the info.  Also clear the
'variable holding message character length.
mnMessageCount = -1
mlMessageChars = 0

'Set the flag to wait for connections to open
'and the initial server reply to arrive
mbGotData = False

'If not connected, then connect to the remote port
If Winsock1.State <> sckConnected Then
    Winsock1.Connect sRemoteAddress, nRemotePort
End If

'Now wait for the server reply
Do Until mbGotData = True
    DoEvents
Loop
ProgressBar1.Value = 20
'Send username
bResult = POP3SendString("USER " & sUserName)
'Check for errors
If bResult = False Then GoTo CheckExitPoint
ProgressBar1.Value = 40
'Send the password
bResult = POP3SendString("PASS " & sPassword)
ProgressBar1.Value = 60
'Check for errors
If bResult = False Then GoTo CheckExitPoint

'Request the Mailbox Status
bResult = POP3SendString("STAT")
'Check for errors
If bResult = False Then GoTo CheckExitPoint
ProgressBar1.Value = 80
'Tell the POP3 Server we are done
bResult = POP3SendString("QUIT")
ProgressBar1.Value = 100
CheckExitPoint:
    'Close the port
    Winsock1.Close
    ProgressBar1.Value = 0
'Display the message info or an error
Select Case mnMessageCount
    Case 0
        MsgBox "There are no messages waiting on the server..."
    
    Case 1
        MsgBox "There is one message waiting on the server. It is " & mlMessageChars & " bytes in size."
        
    Case -1
        MsgBox "An error occured getting message from the server."
        
    Case Else
        'There is more than one message waiting so show the info
        MsgBox "There are " & mnMessageCount & " messages on the server. They are " & mlMessageChars & " total bytes in size."
        
End Select
        
End Sub

Private Function POP3SendString(sCommand As String) As Boolean

Dim sActiveCommand As String
Dim sWorkStr As String
Dim nCharLoc As String

'This routine checks the command to ensure it is the program
'it designed to parse
sActiveCommand = Left$(UCase(Trim$(sCommand)), 4)
Select Case sActiveCommand
    Case "USER", "PASS", "STAT", "QUIT"
    'Valid command, so just display it
    If sActiveCommand = "PASS" Then
    'Don't show the password itself
    lstStatus.AddItem "Server Command: PASS ********"
    Else
    'Otherwise, show the whole command
    lstStatus.AddItem "Server Command: " & sCommand
    End If
    
    Case Else
    'This is a command we are not set up to parse
    MsgBox "Unhandled POP3 command detected: " & sCommand
    Exit Function
    
End Select
    
'Set the flag to wait for the data from the server
mbGotData = False

'Send the string to the POP3 server
Winsock1.SendData sCommand & vbCrLf  'vbCrLf is the Enter button on keyboard

'Wait for the data to get here. Data is stored
'in the msCommand from-scope variable
Do Until mbGotData = True
    DoEvents
    Loop
    
'Parse the data in the reply from the server
'White space is removed and the string is uppercased
'to make the command parsing simpler
Select Case Left$(Trim$(UCase$(msCommand)), 3)
    Case "+OK"  'Command accepted
    'Parse the command specific replies
Select Case sActiveCommand
    Case "USER" 'Name is valid
    Case "PASS"  'Password is accepted
    Case "STAT"
        'Add the reply data to the list box
        lstStatus.AddItem "Server Reply: " & msCommand
        'Parse out the message count and size
        'Start by removing the "+OK " string
        sWorkStr = Right$(msCommand, Len(msCommand) - 4)
        
        'Now find the space that delimits the
        'message count and bytes size data
        nCharLoc = InStr(1, sWorkStr, " ")
        'Now extract the two values
        mnMessageCount = CInt(Mid$(sWorkStr, 1, nCharLoc - 1))
        mlMessageChars = CLng(Mid$(sWorkStr, nCharLoc + 1))
        
    Case "QUIT"  'Ready to disconnect
    
End Select

POP3SendString = True  'Return success

Case "-ER"  'Got an error
    'Add the error info to the list box
    lstStatus.AddItem "Error with command: " & sActiveCommand
    
    POP3SendString = False  'Return failure
    
Case Else    'Unexpected data from the server
    'Update the list box
    lstStatus.AddItem "Unexpected data from POP3 server."
    lstStatus.AddItem "Data: " & msCommand
    
    POP3SendString = False   'Return failure
    
End Select

End Function
