VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{5C4592BE-A01B-11D3-AFAF-BF3F431B043C}#1.0#0"; "TOOLBAR2.OCX"
Begin VB.Form frmFullScreen 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3960
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin AIFCmp1.asxToolbar toolBar 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   873
      FixedSize       =   32
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      ButtonCount     =   14
      PlaySounds      =   0   'False
      CaptionOptions  =   2
      ButtonCaption1  =   "Back"
      ButtonDescription1=   "Display previous page from History"
      ButtonPicture1  =   "frmFullScreen.frx":0000
      ButtonPictureOver1=   "frmFullScreen.frx":0502
      ButtonToolTipText1=   "Back"
      ButtonCaption2  =   "Next"
      ButtonDescription2=   "Display next page from history"
      ButtonPicture2  =   "frmFullScreen.frx":0A04
      ButtonPictureOver2=   "frmFullScreen.frx":0F06
      ButtonToolTipText2=   "Next"
      ButtonCaption3  =   "Stop"
      ButtonDescription3=   "Stop loading a page"
      ButtonPicture3  =   "frmFullScreen.frx":1408
      ButtonPictureOver3=   "frmFullScreen.frx":190A
      ButtonToolTipText3=   "Stop"
      ButtonStyle4    =   0
      ButtonCaption5  =   "Refresh"
      ButtonDescription5=   "Refresh the current page"
      ButtonPicture5  =   "frmFullScreen.frx":1E0C
      ButtonPictureOver5=   "frmFullScreen.frx":230E
      ButtonToolTipText5=   "Refresh"
      ButtonDescription6=   "Displays your home page"
      ButtonPicture6  =   "frmFullScreen.frx":2810
      ButtonPictureOver6=   "frmFullScreen.frx":2D12
      ButtonToolTipText6=   "Home"
      ButtonStyle7    =   2
      ButtonDescription8=   "Displays a search engine"
      ButtonKey8      =   "Search"
      ButtonPicture8  =   "frmFullScreen.frx":3214
      ButtonPictureOver8=   "frmFullScreen.frx":3716
      ButtonToolTipText8=   "Search"
      ButtonDescription9=   "Displays your favourites menu"
      ButtonKey9      =   "Fav"
      ButtonPicture9  =   "frmFullScreen.frx":3C18
      ButtonPictureOver9=   "frmFullScreen.frx":411A
      ButtonToolTipText9=   "Favourites"
      ButtonStyle10   =   2
      ButtonDescription11=   "Allows you to set options"
      ButtonPicture11 =   "frmFullScreen.frx":461C
      ButtonPictureOver11=   "frmFullScreen.frx":4B1E
      ButtonPictureDown11=   "frmFullScreen.frx":5020
      ButtonToolTipText11=   "Options"
      ButtonStyle12   =   2
      ButtonChecked13 =   -1  'True
      ButtonDescription13=   "Displays the page full screen"
      ButtonPicture13 =   "frmFullScreen.frx":5522
      ButtonPictureOver13=   "frmFullScreen.frx":5A24
      ButtonToolTipText13=   "Full Screen"
      ButtonStyle14   =   2
      Begin Prostart.ctlProgress ProgressBar1 
         Height          =   255
         Left            =   6360
         TabIndex        =   3
         Top             =   120
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         Appearance      =   1
         ForeColor       =   0
         BackColor       =   -2147483634
         FillColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionStyle    =   2
         Caption         =   ""
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "GO"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   290
         Left            =   7600
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4800
         TabIndex        =   1
         Text            =   "http://"
         Top             =   120
         Width           =   2800
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   4695
      ExtentX         =   8281
      ExtentY         =   6165
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Menu mnuBack 
      Caption         =   "back"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuForward 
      Caption         =   "forward"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmFullScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
WebBrowser1.Navigate Text1.Text
End Sub

Private Sub Form_Load()
WebBrowser1.Navigate frmMain.WebBrowser1.LocationURL
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Text1.ForeColor = vbBlack
Text1.FontUnderline = False
End Sub

Private Sub Form_Resize()
ProgressBar1.left = frmFullScreen.Width - 2600
WebBrowser1.Width = frmFullScreen.Width

End Sub

Private Sub mnuBack_Click()
On Error Resume Next
WebBrowser1.GoBack
End Sub


Private Sub mnuForward_Click()
On Error Resume Next
WebBrowser1.GoForward
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Text1.ForeColor = vbBlue
Text1.FontUnderline = True
End Sub

Private Sub toolBar_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
If ButtonIndex = 1 Then mnuBack_Click: If ButtonIndex = 2 Then mnuForward_Click
If ButtonIndex = 3 Then WebBrowser1.Stop

If ButtonIndex = 5 Then WebBrowser1.Refresh             'For some reason (i don't know why) these
If ButtonIndex = 6 Then WebBrowser1.Navigate homePage   'buttonclicks could not refer to the menu_click event
If ButtonIndex = 8 Then WebBrowser1.GoSearch

If ButtonIndex = 9 Then PopupMenu frmMain.mnuFavourites, 2, , toolBar.top + toolBar.Height

If ButtonIndex = 11 Then frmOptions.Show

If ButtonIndex = 13 Then
Unload Me
frmMain.Show

End If
End Sub


Private Sub toolBar_Resize(ByVal NewWidth As Single, ByVal NewHeight As Single)
WebBrowser1.Width = Me.ScaleWidth
WebBrowser1.Height = Me.ScaleHeight - toolBar.Height
End Sub

Private Sub WebBrowser1_DownloadBegin()
ProgressBar1.Visible = True
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
Text1.Text = WebBrowser1.LocationURL
End Sub

Private Sub WebBrowser1_TitleChange(ByVal Text As String)
frmMain.WebBrowser1.Navigate frmFullScreen.WebBrowser1.LocationURL

End Sub
Private Sub WebBrowser1_DownloadComplete()
ProgressBar1.value = 100
ProgressBar1.Visible = False

End Sub
Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error GoTo progressERR
If Progress = -1 Then ProgressBar1.value = 100
If Progress > 0 And ProgressMax > 0 Then
    ProgressBar1.value = Progress * 100 / ProgressMax
    StatusBar1.Panels(2).Text = ProgressBar1.value & " %"
    End If
    Exit Sub
progressERR:
End Sub
