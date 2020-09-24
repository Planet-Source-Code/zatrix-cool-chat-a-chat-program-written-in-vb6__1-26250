VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      BackColor       =   &H00FF8080&
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4380
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00FF8080&
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         MaskColor       =   &H00FF8080&
         TabIndex        =   7
         Top             =   3480
         Width           =   975
      End
      Begin VB.PictureBox picLogo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   360
         Picture         =   "About.frx":0000
         ScaleHeight     =   585
         ScaleWidth      =   615
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "This product is FREEWARE !"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Tag             =   "LicenseTo"
         Top             =   1800
         Width           =   3135
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   720
         TabIndex        =   6
         Tag             =   "Product"
         Top             =   360
         Width           =   3435
      End
      Begin VB.Label lblCompanyProduct 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "MDSoftware"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   435
         Left            =   720
         TabIndex        =   5
         Tag             =   "CompanyProduct"
         Top             =   1200
         Width           =   3360
      End
      Begin VB.Label lblPlatform 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Win95/98"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   360
         Left            =   1620
         TabIndex        =   4
         Tag             =   "Platform"
         Top             =   2880
         Width           =   1305
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   1680
         TabIndex        =   3
         Tag             =   "Version"
         Top             =   840
         Width           =   1290
      End
      Begin VB.Label lblCompany 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "ZATRiX"
         Height          =   255
         Left            =   1680
         TabIndex        =   2
         Tag             =   "Company"
         Top             =   2400
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

