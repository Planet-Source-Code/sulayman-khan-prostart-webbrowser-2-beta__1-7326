VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prostart Options (Currently Working On)"
   ClientHeight    =   4680
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5535
      Begin VB.TextBox Text5 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   23
         Top             =   3240
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2760
         TabIndex        =   22
         Top             =   2760
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2760
         TabIndex        =   17
         Text            =   "mail.yourmailserver.com"
         Top             =   2280
         Width           =   2655
      End
      Begin VB.CommandButton cmdInternet 
         Caption         =   "Internet Propeties"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2760
         TabIndex        =   15
         Text            =   "you@yourdomain.com"
         Top             =   1560
         Width           =   2655
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Disable Popups"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3120
         TabIndex        =   13
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   10
         Text            =   "Not Implemented Yet"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "Password:"
         Height          =   255
         Left            =   1080
         TabIndex        =   24
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Username:"
         Height          =   255
         Left            =   1080
         TabIndex        =   21
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Mail Server:"
         Height          =   255
         Left            =   1080
         TabIndex        =   20
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Recieving Mail"
         ForeColor       =   &H8000000A&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1920
         Width           =   5295
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sending Mail"
         ForeColor       =   &H8000000A&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   5295
      End
      Begin VB.Label Label3 
         Caption         =   "Your Replying E-mail:"
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Browser Popup Window:"
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Set your homepage:"
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmOptions.frx":000C
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   2100
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   1545
         TabIndex        =   7
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   645
         TabIndex        =   6
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   4320
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dblReturn As Double
Private Sub cmdApply_Click()
    frmSendMail.txtFromEmail = Me.Text2
    frmGetMail.tServer = Me.Text3
    frmGetMail.tName = Me.Text4
    frmGetMail.tPassword = Me.Text5
    MsgBox "You have now changed the settings of your browser"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdInternet_Click()
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0", 5)
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub


Private Sub Form_Load()
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Check1.value = frmMain.Check1.value
    Check1.Caption = frmMain.Check1.Caption
End Sub

