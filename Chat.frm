VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmChat 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cool Chat"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   7155
   FillColor       =   &H00008080&
   HelpContextID   =   10
   Icon            =   "Chat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   5805
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   2469
            MinWidth        =   2469
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   5115
            MinWidth        =   5115
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "12:33 AM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSendText 
      Caption         =   "Send Text"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      ToolTipText     =   "This button send the text to the left to the other party"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtSendText 
      Height          =   285
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "Text to send to your friend"
      Top             =   1920
      Width           =   4335
   End
   Begin VB.TextBox txtIncomingData 
      Height          =   1935
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
   Begin VB.Frame FrameSendEmail 
      BackColor       =   &H00FF8080&
      Caption         =   "Send Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   3135
      Left            =   120
      TabIndex        =   23
      Top             =   2400
      Visible         =   0   'False
      Width           =   5535
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "CLOSE"
         Height          =   255
         Index           =   1
         Left            =   4320
         TabIndex        =   32
         Top             =   2280
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "SEND"
         Height          =   255
         Index           =   0
         Left            =   4320
         TabIndex        =   31
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FF8080&
         Caption         =   "Status"
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
         Left            =   120
         TabIndex        =   28
         Top             =   2160
         Width           =   3975
         Begin ComctlLib.ProgressBar ProgressBar1 
            Height          =   135
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   238
            _Version        =   327682
            Appearance      =   1
         End
         Begin VB.Label StatusTxt 
            BackColor       =   &H00FF8080&
            Height          =   255
            Left            =   120
            TabIndex        =   30
            ToolTipText     =   "Status Messages"
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.TextBox txtEmailBodyOfMessage 
         Height          =   1335
         Left            =   2040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         ToolTipText     =   "Put your text here"
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox txtToEmailAddress 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtFromEmailAddress 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "youremail@yourisp.com"
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtEmailServer 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "12.9.25.36"
         ToolTipText     =   "mail.yourisp.com or smtp.yourisp.com"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txtEmailSubject 
         Height          =   285
         Left            =   2880
         TabIndex        =   4
         ToolTipText     =   "Email Subject"
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF8080&
         Caption         =   "Send To:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "Your Email Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF8080&
         Caption         =   "SMTP Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF8080&
         Caption         =   "Subject:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   24
         Top             =   1680
         Width           =   855
      End
   End
   Begin VB.Frame FrameOptions 
      BackColor       =   &H00FF8080&
      Caption         =   "Connect Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   1455
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   3495
      Begin VB.Timer tmrFlash 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2040
         Top             =   240
      End
      Begin VB.Timer tmrClock 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3000
         Top             =   600
      End
      Begin VB.TextBox txtNick 
         Height          =   285
         Left            =   2160
         TabIndex        =   12
         Text            =   "Nick Name"
         ToolTipText     =   "You can change the Nickname even during a chat session"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Text            =   "1005"
         ToolTipText     =   "The Host AND Guest MUST use the same port number"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Text            =   "localhost"
         ToolTipText     =   "If you're connecting to yourself, then leave as Localhost... Otherwise enter the IP# of the Host"
         Top             =   1080
         Width           =   1415
      End
      Begin VB.OptionButton optHostGuest 
         BackColor       =   &H00FF8080&
         Caption         =   "Guest"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   9
         ToolTipText     =   "You will connect to a host... Get the IP# of the Host and enter below"
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optHostGuest 
         BackColor       =   &H00FF8080&
         Caption         =   "Host"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   8
         ToolTipText     =   "You will receive the connection, tell your friend your IP#"
         Top             =   360
         Width           =   735
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   3000
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock Winsock2 
         Left            =   2520
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label LabelAddress 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   690
      End
      Begin VB.Label labelPort 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1635
         TabIndex        =   15
         Top             =   840
         Width           =   375
      End
   End
   Begin VB.Frame FrameCounter 
      BackColor       =   &H00FF8080&
      Caption         =   "Online For:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   18
      Top             =   3960
      Width           =   2655
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Seconds"
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
         Left            =   1680
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblSeconds 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   1320
         TabIndex        =   21
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Minute(s)"
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
         Left            =   480
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblMinutes 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      ToolTipText     =   "This button controls Connecting, Disconnecting and Listening"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Image imgConnecting 
      Height          =   1350
      Left            =   3720
      Picture         =   "Chat.frx":27A2
      Top             =   2520
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Image imgOffline 
      Height          =   1350
      Left            =   3720
      Picture         =   "Chat.frx":379B
      Top             =   2520
      Width           =   1110
   End
   Begin VB.Image imgOnline 
      Height          =   1350
      Left            =   3720
      Picture         =   "Chat.frx":4793
      Top             =   2520
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuUtilities 
      Caption         =   "&Utilites"
      Begin VB.Menu mnuSendEmail 
         Caption         =   "Send Email"
      End
      Begin VB.Menu mnuCheckEmail 
         Caption         =   "Check Email"
      End
   End
   Begin VB.Menu mnuClearIncoming 
      Caption         =   "&Clear Screen"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Contents"
         HelpContextID   =   10
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Notes:
'In case you can't figure out how I got all the images
'to change, depending on the connection status, I have 4
'different images laid on top of each other (imgConnecting,
'imgOnline, imgOffline, imgConnecting (you can move
'them around on the form layout to see))

'Chat variables
Dim IP As String
Dim s As Integer
Dim m As Integer

'Utilities: Send Email variables
Dim Response As String, Reply As Integer, DateNow As String
Dim first As String, Second As String, Third As String
Dim Fourth As String, Fifth As String, Sixth As String
Dim Seventh As String, Eighth As String, Ninth As String
Dim start As Single, Tmr As Single

'################# BEGIN CHAT SUBS ######################

'Below Subs (Offline, Online, Connecting & Listening
'are repetitious tasks that are called from other
'Subs. Instead of putting these 10 or so lines in each Sub
'that they belong to, I put them up here which makes it
'easier to figure out the heart of each sub instead of
'being bombarded by all of these .Visible's ,  etc
'Perhaps I could've made some modules?

Public Sub Offline()
imgOffline.Visible = True
imgOnline.Visible = False
imgConnecting.Visible = False
txtPort.Enabled = True
txtNick.Enabled = True
txtIP.Enabled = True
cmdSendText.Enabled = False
txtSendText.Enabled = False
optHostGuest(0).Enabled = True
optHostGuest(1).Enabled = True
StatusBar1.Panels(2).Text = ""
Winsock1.Close

'Reset Online Timer
tmrClock.Enabled = False
s = 0
m = 0
FrameCounter.Visible = False
lblSeconds.Caption = ""
lblMinutes.Caption = ""

End Sub

Public Sub Online()
imgOnline.Visible = True
imgOffline.Visible = False
imgConnecting.Visible = False
txtIP.Enabled = False
txtPort.Enabled = False
cmdSendText.Enabled = True
txtSendText.Enabled = True
txtIncomingData.Text = ""
optHostGuest(0).Enabled = False
optHostGuest(1).Enabled = False
cmdConnect.Caption = "Disconnect"

End Sub

Public Sub Connecting()
imgConnecting.Visible = True
imgOffline.Visible = False
imgOnline.Visible = False
cmdSendText.Enabled = False
txtSendText.Enabled = False
StatusBar1.Panels(2).Text = "CONTACTING HOST..."
If Winsock1.State <> sckClosed Then
Winsock1.Close
End If
           IP = txtIP.Text
If LCase$(IP) = "localhost" Then IP = Winsock1.LocalIP
Winsock1.Connect txtIP.Text, txtPort.Text

End Sub

Public Sub Listening()
imgConnecting.Visible = True
imgOffline.Visible = False
imgOnline.Visible = False
cmdSendText.Enabled = False
txtSendText.Enabled = False

Winsock1.Close
Winsock1.LocalPort = txtPort.Text 'set the port
Winsock1.Listen 'tell it to listen
StatusBar1.Panels(2).Text = "LISTENING ON PORT: " & txtPort.Text


End Sub

Private Sub cmdConnect_Click()
    If cmdConnect.Caption = "Connect" Or cmdConnect.Caption = "Listen" Then
'If the button is showing Connect or Listen, then we are
'currently offline, so do this:

'Now do the Online sub
    Call Online

    Else
'If the button is NOT showing Connect or Listen, then we
'must be Offline, so call the Offline sub
    Call Offline
       
     If optHostGuest(0).Value = True Then
cmdConnect.Caption = "Listen"
   
'If the Host button is checked, then we dont want
'to do the Online sub, which is for Guest Mode, we
'want to call the Listening Sub:

   Call Listening
       Else
'Else we are already Listening in Host mode, so Go Offline
cmdConnect.Caption = "Connect"
       End If
Call Offline

Exit Sub
    End If
  
    Select Case optHostGuest(0).Value
        Case True:  'Host
           'Listen for connections
Call Listening
        Case False: 'Guest
           'Try to connect
Call Connecting
    End Select

End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuCheckEmail_Click()
frmCheckEmail.Show
End Sub

Private Sub mnuClearIncoming_Click()
txtIncomingData.Text = "" 'Clear chat screen
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuHelpContents_Click()
Shell "winhelp.exe coolchat.hlp", vbNormalFocus

End Sub
Private Sub mnuSendEmail_Click()
FrameOptions.Visible = False
FrameSendEmail.Visible = True
FrameCounter.Visible = False
IP = Winsock1.LocalIP
txtEmailBodyOfMessage.Text = "My IP number is: " & IP
End Sub

Private Sub optHostGuest_Click(Index As Integer)
'Does something when either Guest or Host value is checked
'(Not pushing the connect button, just clicking the options)
    
    Select Case Index
       
       Case 0: 'Host value is clicked on
            'Automatically paste your IP# in the Address box
           
          IP = Winsock1.LocalIP
          txtIP.Text = IP
           'Prevent the address from being changed
           txtIP.Locked = True
          txtIP.Enabled = False
          cmdConnect.Caption = "Listen"
        StatusBar1.Panels(1).Text = "HOST MODE"
       
       Case 1: 'Guest value is clicked on
              txtIP.Text = "localhost"
          cmdConnect.Caption = "Connect"
       txtIP.Locked = False
        txtIP.Enabled = True
        StatusBar1.Panels(1).Text = "GUEST MODE"
          
    End Select
End Sub

Private Sub cmdSendText_Click()
'Send Text to other party
Winsock1.SendData txtNick.Text & ":  " & txtSendText.Text
'Display YOUR text on your screen, too
txtIncomingData.Text = txtIncomingData.Text + vbCrLf + txtNick.Text + ": " + txtSendText.Text
'Make the text box blank after we send the text.
'otherwise, we'd have to erase all the characters before
'sending a new message, EVERYTIME
txtSendText.Text = ""
End Sub

Private Sub Form_Load()
        StatusBar1.Panels(1).Text = "GUEST MODE"
cmdSendText.Enabled = False
txtSendText.Enabled = False
FrameCounter.Visible = False
StatusBar1.Panels(2).Text = "Cool Chat v" & App.Major & "." & App.Minor & "." & App.Revision & "  - MDSoftware"
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0:
On Error GoTo ErrorHandling
    SendEmail txtEmailServer.Text, txtFromEmailAddress.Text, txtFromEmailAddress.Text, txtToEmailAddress.Text, txtToEmailAddress.Text, txtEmailSubject.Text, txtEmailBodyOfMessage.Text
 'Make the Send Mail radio button be blank again
 Option1(0).Value = False
 MsgBox ("Mail Sent")
    StatusTxt.Caption = "Mail Sent"
    StatusTxt.Refresh
    Beep

    Close

ErrorHandling:
Winsock2.Close

Case 1:
FrameOptions.Visible = True
FrameSendEmail.Visible = False
Option1(1).Value = False
If imgOnline.Visible = True Then
FrameCounter.Visible = True
End If
End Select
End Sub

Private Sub tmrClock_Timer()

'This does nothing but simply display an online timer
'when a connection is complete
'Note if you want to make a timer like this, make sure
'you set the Interval in Properties to 1000
'Simplest form to do a timer is:
's = s + 1
'Label1.Caption = s (for Label) or Text1.Text = s (for text)
'Just those 2 lines will make you a seconds counter.

s = s + 1
'When seconds gets to 60, then restart to 0 and add 1 to minutes
If s = 60 Then
s = 0
m = m + 1
End If

'When the first minutes is reached display it
If m > 0 Then
lblMinutes.Caption = m
End If
lblSeconds.Caption = s
DoEvents

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key ' start select method
Case "One"
cmdConnect_Click
Case "Three"
txtIncomingData.Text = ""
Case "Five"
frmSendEmail.Show
Case "Six"
frmCheckEmail.Show

End Select ' end select method

End Sub

Private Sub txtIncomingData_Change()
'Makes the chat screen auto scroll
'BTW, Len is used to tell you the length (# of characters)
'in a text ("string") Variable
txtIncomingData.SelStart = Len(txtIncomingData)
End Sub

Private Sub txtSendText_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then 'If user pressed 'Enter'
      cmdSendText_Click 'click 'Send' button
      KeyAscii = 0 'Make sure it doesnt write enter to txtText
   End If
    
End Sub


Private Sub Winsock1_Close()
'What happens when a connection is closed/terminated

    Select Case optHostGuest(0).Value
        Case True:  'Host closed Winsock
cmdConnect_Click    'Click the Disconnect button
cmdConnect_Click     'Click Listen button again for Listen Mode

MsgBox "Connection terminated by Guest. Server has been reset and awaiting a new client..."

        Case False: 'Guest closed Winsock


MsgBox "Connection terminated by Host..."

cmdConnect_Click     'Push the disconnect button to make offline
                    

    End Select

End Sub

Private Sub Winsock1_Connect()

    Select Case optHostGuest(0).Value
        
        Case True:  'Host got a connection

'I left this blank because the "CLIENT CONNECTED" msg
'is displayed below in the Winsock1_ConnectionRequest Sub
'along with the Call Online command

        Case False: 'Guest got a connection

Call Online
StatusBar1.Panels(2).Text = "CONNECTED TO HOST"
End Select

'Turn on and show Online Timer
FrameCounter.Visible = True
tmrClock.Enabled = True

End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)

'When Host received a Connection Request...

If Winsock1.State <> sckClosed Then
    Winsock1.Close
Winsock1.LocalPort = txtPort.Text
Winsock1.Accept requestID 'accept the connection
Online
StatusBar1.Panels(2).Text = "CLIENT CONNECTED"

FrameCounter.Visible = True
tmrClock.Enabled = True

End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

'Get incoming chat data and put it in the Text Box
Dim Data As String
Winsock1.GetData Data 'gets the data
txtIncomingData.Text = txtIncomingData.Text + vbCrLf & Data
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'What to do if an error occurs, this usually happens
'the Host is not Listening or your ISP is down
StatusBar1.Panels(2).Text = "WINSOCK ERROR: " & Err
cmdConnect_Click   'Push the disconnect button to reset

MsgBox "UNABLE TO CONNECT..."

End Sub

'#################### END CHAT SUBS ###############



'#################### SEND EMAIL SUBS #######################

Sub SendEmail(MailServerName As String, FromName As String, FromEmailAddress As String, ToName As String, ToEmailAddress As String, EmailSubject As String, EmailBodyOfMessage As String)
  On Error GoTo ErrorHandler
  
    Winsock2.LocalPort = 0 ' Must set local port to 0 (Zero) or you can only send 1 e-mail pre program start
    
If Winsock2.State = sckClosed Then ' Check to see if socet is closed
    DateNow = Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(Time, "hh:mm:ss") & "" & " -0600"
    first = "mail from:" + Chr(32) + FromEmailAddress + vbCrLf ' Get who's sending E-Mail address
    Second = "rcpt to:" + Chr(32) + ToEmailAddress + vbCrLf ' Get who mail is going to
    Third = "Date:" + Chr(32) + DateNow + vbCrLf ' Date when being sent
    Fourth = "From:" + Chr(32) + FromName + vbCrLf ' Who's Sending
    Fifth = "To:" + Chr(32) + ToEmailAddress + vbCrLf ' Who it going to
    Sixth = "Subject:" + Chr(32) + EmailSubject + vbCrLf ' Subject of E-Mail
    Seventh = EmailBodyOfMessage + vbCrLf ' E-mail message body
    Ninth = "X-Mailer: EBT Reporter v 2.x" + vbCrLf ' What program sent the e-mail, customize this
    Eighth = Fourth + Third + Ninth + Fifth + Sixth  ' Combine for proper SMTP sending

    Winsock2.Protocol = sckTCPProtocol ' Set protocol for sending
    Winsock2.RemoteHost = MailServerName ' Set the server address
    Winsock2.RemotePort = 25 ' Set the SMTP Port
    Winsock2.Connect ' Start connection
   
    WaitFor ("220")
    
    StatusTxt.Caption = "Connecting...."
    StatusTxt.Refresh
    Winsock2.SendData ("HELO worldcomputers.com" + vbCrLf)

    WaitFor ("250")

    StatusTxt.Caption = "Connected"
    StatusTxt.Refresh
 ProgressBar1.Value = 25
    Winsock2.SendData (first)

    StatusTxt.Caption = "Sending Message"
    StatusTxt.Refresh

    WaitFor ("250")

    Winsock2.SendData (Second)

    WaitFor ("250")
 ProgressBar1.Value = 50

    Winsock2.SendData ("data" + vbCrLf)
    
    WaitFor ("354")


    Winsock2.SendData (Eighth + vbCrLf)
    Winsock2.SendData (Seventh + vbCrLf)
    Winsock2.SendData ("." + vbCrLf)

    WaitFor ("250")
 ProgressBar1.Value = 75
    Winsock2.SendData ("quit" + vbCrLf)
     ProgressBar1.Value = 100
    StatusTxt.Caption = "Disconnecting"
    StatusTxt.Refresh

    WaitFor ("221")

    Winsock2.Close
     ProgressBar1.Value = 0
Else
    MsgBox (Str(Winsock2.State))
End If
ErrorHandler:
 ProgressBar1.Value = 0
 StatusTxt.Caption = "Winsock Error: " & Err
End Sub
Sub WaitFor(ResponseCode As String)
    start = Timer ' Time event so won't get stuck in loop
    While Len(Response) = 0
        Tmr = start - Timer
        DoEvents ' Let System keep checking for incoming response **IMPORTANT**
        If Tmr > 50 Then ' Time in seconds to wait
            MsgBox "SMTP service error, timed out while waiting for response", 64
            Exit Sub
        End If
    Wend
    While Left(Response, 3) <> ResponseCode
        DoEvents
        If Tmr > 50 Then
            MsgBox "SMTP service error, impromper response code. Code should have been: " + ResponseCode + " Code recieved: " + Response, 64
            Exit Sub
        End If
    Wend
Response = "" ' Sent response code to blank **IMPORTANT**
End Sub

Private Sub Form_Unload(cancel As Integer)
Winsock2.Close
End Sub


Private Sub winsock2_DataArrival(ByVal bytesTotal As Long)

    Winsock2.GetData Response ' Check for incoming response *IMPORTANT*

End Sub
