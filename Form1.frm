VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{5C4592BE-A01B-11D3-AFAF-BF3F431B043C}#1.0#0"; "TOOLBAR2.OCX"
Begin VB.Form frmMain 
   Caption         =   "Prostart Web Browser"
   ClientHeight    =   7395
   ClientLeft      =   960
   ClientTop       =   855
   ClientWidth     =   10050
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   10050
   WindowState     =   2  'Maximized
   Begin Prostart.ctlProgress ProgressBar1 
      Height          =   220
      Left            =   0
      TabIndex        =   5
      Top             =   7170
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   397
      Appearance      =   1
      ForeColor       =   0
      BackColor       =   -2147483634
      FillColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ProgressBar pBar 
      Height          =   220
      Left            =   0
      TabIndex        =   4
      Top             =   8520
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   397
      _Version        =   327682
      Appearance      =   1
      Enabled         =   0   'False
   End
   Begin VB.ComboBox cboAddress 
      Height          =   315
      ItemData        =   "Form1.frx":030A
      Left            =   0
      List            =   "Form1.frx":030C
      TabIndex        =   3
      Top             =   480
      Width           =   10695
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      ItemData        =   "Form1.frx":030E
      Left            =   0
      List            =   "Form1.frx":0315
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   8160
      Visible         =   0   'False
      Width           =   3495
   End
   Begin AIFCmp1.asxToolbar toolBar 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   873
      FixedSize       =   32
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      ButtonCount     =   13
      CaptionOptions  =   2
      HotTracking     =   -1  'True
      ButtonCaption1  =   "Back"
      ButtonDescription1=   "Display previous page from History"
      ButtonPicture1  =   "Form1.frx":034C
      ButtonPictureOver1=   "Form1.frx":084E
      ButtonToolTipText1=   "Back"
      ButtonCaption2  =   "Next"
      ButtonDescription2=   "Display next page from history"
      ButtonPicture2  =   "Form1.frx":0D50
      ButtonPictureOver2=   "Form1.frx":1252
      ButtonToolTipText2=   "Next"
      ButtonCaption3  =   "Stop"
      ButtonDescription3=   "Stop loading a page"
      ButtonPicture3  =   "Form1.frx":1754
      ButtonPictureOver3=   "Form1.frx":1C56
      ButtonToolTipText3=   "Stop"
      ButtonStyle4    =   0
      ButtonCaption5  =   "Refresh"
      ButtonDescription5=   "Refresh the current page"
      ButtonPicture5  =   "Form1.frx":2158
      ButtonPictureOver5=   "Form1.frx":265A
      ButtonToolTipText5=   "Refresh"
      ButtonDescription6=   "Displays your home page"
      ButtonPicture6  =   "Form1.frx":2B5C
      ButtonPictureOver6=   "Form1.frx":305E
      ButtonToolTipText6=   "Home"
      ButtonStyle7    =   2
      ButtonDescription8=   "Displays a search engine"
      ButtonKey8      =   "Search"
      ButtonPicture8  =   "Form1.frx":3560
      ButtonPictureOver8=   "Form1.frx":3A62
      ButtonToolTipText8=   "Search"
      ButtonDescription9=   "Displays your favourites menu"
      ButtonKey9      =   "Fav"
      ButtonPicture9  =   "Form1.frx":3F64
      ButtonPictureOver9=   "Form1.frx":4466
      ButtonToolTipText9=   "Favourites"
      ButtonStyle10   =   2
      ButtonDescription11=   "Allows you to set options"
      ButtonPicture11 =   "Form1.frx":4968
      ButtonPictureOver11=   "Form1.frx":4E6A
      ButtonPictureDown11=   "Form1.frx":536C
      ButtonToolTipText11=   "Options"
      ButtonStyle12   =   2
      ButtonDescription13=   "Displays the page full screen"
      ButtonPicture13 =   "Form1.frx":586E
      ButtonPictureOver13=   "Form1.frx":5D70
      ButtonToolTipText13=   "Full Screen"
      Begin VB.CheckBox Check1 
         Caption         =   "Disable Popups"
         Height          =   450
         Left            =   4800
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Disables/Enables the popups of your browser"
         Top             =   25
         Width           =   700
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   7125
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3828
            MinWidth        =   3828
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   13432
            MinWidth        =   1182
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   10695
      ExtentX         =   18865
      ExtentY         =   13361
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
   Begin ComctlLib.ProgressBar ProgressBar 
      Height          =   242
      Left            =   20
      TabIndex        =   6
      Top             =   7150
      Width           =   2185
      _ExtentX        =   3863
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New window..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpenExternal 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu div1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu divider1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "C&opy"
         Shortcut        =   ^C
      End
      Begin VB.Menu divider2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuGoTo 
         Caption         =   "&Go to"
         Begin VB.Menu mnuBack 
            Caption         =   "&Back"
         End
         Begin VB.Menu mnuForward 
            Caption         =   "&Forward"
         End
         Begin VB.Menu div 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHome 
            Caption         =   "&Home Page"
         End
         Begin VB.Menu mnuSearch 
            Caption         =   "&Search"
         End
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Sto&p"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuSource 
         Caption         =   "So&urce"
         Shortcut        =   ^S
      End
      Begin VB.Menu div2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFullScreen 
         Caption         =   "Full &Screen!"
      End
   End
   Begin VB.Menu mnuFavourites 
      Caption         =   "F&avourites"
      Begin VB.Menu mnuAddFav 
         Caption         =   "&Add to favourites..."
      End
      Begin VB.Menu mnuOrgFav 
         Caption         =   "&Organize favourites..."
      End
      Begin VB.Menu mnuArray 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuOptionB 
         Caption         =   "&Browser Properties"
      End
      Begin VB.Menu div6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMail 
         Caption         =   "Mail"
         Begin VB.Menu mnuMailSend 
            Caption         =   "&Send Mail"
         End
         Begin VB.Menu mnuRecieve 
            Caption         =   "&Recieve Mail"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu div3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProstartWeb 
         Caption         =   "Prostart on the &web"
         Begin VB.Menu mnuUpdate 
            Caption         =   "&Check for updates"
         End
         Begin VB.Menu mnuProstartHome 
            Caption         =   "&Prostart homepage"
         End
      End
      Begin VB.Menu div4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Used for menu-favourites
Public DaBa As String, CurNew As String, NewNod As Integer
Private File As String

'Used for menu-favourites

Private initialDir As String ' Used when opening local files
Const homePage = "www.microsoft.com" ' Set home URL

Private Sub Command1_Click()
WebBrowser1.Navigate Text1.Text
End Sub

Private Sub cboAddress_Click()
WebBrowser1.Navigate cboAddress.Text 'User choose a file in 'history'
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer) 'This event logs all visited webpages in the address combo

If KeyAscii = 13 Then                       'If user press enter button
WebBrowser1.Navigate cboAddress.Text        'Go to page
comboLoop                                   'See sub below!
End If
End Sub


Sub comboLoop() 'A bad name - i know - what it does is add the address to the list if it's not there from before
Dim inList As Boolean
inList = False                              'Set inList as false initially

For i = 0 To cboAddress.ListCount           'Loop through the combo to check if address exist from before
  If cboAddress.Text = cboAddress.List(i) Then inList = True    'If it's already there, set boolean to true
Next i

If inList = False Then cboAddress.AddItem cboAddress.Text       'If, after the loop - inlist is still false
                                                                'we can assume that the address did not exist
                                                                'in the combo - so we add it!
End Sub

Private Sub Check1_Click()
If Check1.value = 1 Then
    Check1.Caption = "Enable Popups"
Else
    Check1.Caption = "Disable Popups"
End If
End Sub

Private Sub Form_Load()
    File = App.Path & "\favourites.rtx" 'This is an ini file, but let's not give that away... ;)
    initialDir = App.Path & "\"
    
    LoadINI
    extractMenu
    
    WebBrowser1.Navigate homePage

    If numberOfWindows = Empty Then numberOfWindows = 1 'This variable is set in module1 and we use it to
                                                        'keep track of how many windows are opened. If this
                                                        'is the first window - set the counter to 1
End Sub

Private Sub Form_Resize()
On Error Resume Next
WebBrowser1.Width = Me.ScaleWidth
WebBrowser1.Height = Me.ScaleHeight - StatusBar1.Height - toolBar.Height - cboAddress.Height
picBar.top = Me.Height - 930
cboAddress.Width = Me.ScaleWidth
ProgressBar1.top = Me.Height - 930
End Sub

Private Sub Form_Unload(Cancel As Integer) 'Make sure nothing is in memory on exit
Unload frmOpen
Unload frmSplash
Unload frmFullScreen
'numberOfWindows = numberOfWindows - 1 'Update the variable we use to keep track of instances

If numberOfWindows < 1 Then 'If this is the last instance (window) then end
End
End If
End Sub


Private Sub mnuAbout_Click()
frmCredits.Show 1
End Sub

Private Sub mnuAddFav_Click()
Dim locname As String
Dim locurl As String
locname = WebBrowser1.LocationName & "="
locurl = WebBrowser1.LocationURL
End Sub

Private Sub mnuArray_Click(Index As Integer)
WebBrowser1.Navigate GetValue("Favourites", mnuArray.Item(Index).Caption, File)
End Sub

Sub mnuBack_Click()
On Error Resume Next
WebBrowser1.GoBack
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Sub mnuForward_Click()
On Error Resume Next
WebBrowser1.GoForward
End Sub

Private Sub mnuFullScreen_Click()
toolBar.ForceClick (13)
End Sub

Sub mnuHome_Click()
WebBrowser1.Navigate homePage
End Sub



Private Sub mnuMailSend_Click()
frmSendMail.Show
End Sub

Sub mnuNew_Click()
Dim newInstance As New frmMain 'Create a new instance of our form
newInstance.Show
newInstance.Caption = "Prostart Web Browser - New Window"
'Add one to the number of open windows
End Sub

Private Sub mnuOpenExternal_Click()
frmOpen.Show
End Sub


Private Sub mnuOptionB_Click()
frmOptions.Show
End Sub



Private Sub mnuOrgFav_Click()
MsgBox "Not implemented yet"
End Sub

Private Sub mnuRecieve_Click()
frmGetMail.Show
End Sub

Sub mnuRefresh_Click()
WebBrowser1.Refresh
End Sub

Sub mnuSearch_Click()
WebBrowser1.GoSearch
End Sub

Private Sub mnuSource_Click()
frmSource.Show
End Sub

Sub mnuStop_Click()
WebBrowser1.Stop
End Sub

Private Sub toolBar_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
If ButtonIndex = 1 Then mnuBack_Click: If ButtonIndex = 2 Then mnuForward_Click
If ButtonIndex = 3 Then mnuStop_Click: pBar.value = 100

If ButtonIndex = 5 Then WebBrowser1.Refresh             'For some reason (i don't know why) these
If ButtonIndex = 6 Then WebBrowser1.Navigate homePage   'buttonclicks could not refer to the menu_click event
If ButtonIndex = 8 Then WebBrowser1.GoSearch

If ButtonIndex = 9 Then PopupMenu mnuFavourites, 2, , toolBar.top + toolBar.Height

If ButtonIndex = 11 Then frmOptions.Show

If ButtonIndex = 13 Then
Me.Hide
frmFullScreen.Show
End If
End Sub

Private Sub WebBrowser1_DownloadBegin()
ProgressBar1.Visible = True
End Sub

Private Sub WebBrowser1_DownloadComplete()
ProgressBar1.value = 100
ProgressBar1.Visible = False
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean) 'NB!!!     I changed this so the new window
If Check1.value = 1 Then
    Cancel = True
Else
    Cancel = False
End If
End Sub


Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    frmMain.Caption = WebBrowser1.LocationName
        StatusBar1.Panels(2).Text = WebBrowser1.LocationURL
        StatusBar1.ToolTipText = WebBrowser1.LocationURL
End Sub


Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error GoTo progressERR
If Progress = -1 Then ProgressBar1.value = 100
If Progress > 0 And ProgressMax > 0 Then
    ProgressBar1.value = Progress * 100 / ProgressMax
    End If
    Exit Sub
progressERR:
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
StatusBar1.Panels(2).Text = Text
StatusBar1.ToolTipText = WebBrowser1.LocationURL
frmMain.Caption = WebBrowser1.LocationName & " - Prostart Browser"
cboAddress.Text = WebBrowser1.LocationURL
End Sub





'///////The following subs are used in the form load event to create and fill inn favourites in the menu


Sub LoadINI()
DaBa = ""
Dim x, y, Z, GenKey As Integer, A, b, CurInfo As String, CurData As String, CurDir As String, CurDirPos
List1.Clear
DaBa = String(FileLen(File), " ")
Open File For Binary As #1
Get #1, 1, DaBa
Close #1

For y = 1 To Len(DaBa)
    If Mid(DaBa, y, 1) = "[" Then
    For x = y To Len(DaBa)
    
    If Mid(DaBa, x, 1) = "]" Then
    CurDirPos = y + 1
    CurDir = Mid(DaBa, y + 1, (x - y) - 1)
    On Error Resume Next
    'List1.AddItem CurDir
 
           For Z = x + 1 To Len(DaBa)
           If Mid(DaBa, Z, 1) = "[" Then Exit For
           If Mid(DaBa, Z, 1) = "=" Then
               For A = Z To 1 Step -1
               If Mid(DaBa, A, 1) = "]" Then Exit For
                If Mid(DaBa, A, 1) = Chr(13) Then
                CurInfo = Mid(DaBa, A + 2, Z - A - 2)
                List1.AddItem CurInfo
                Exit For
                End If
               Next A
               
               For A = Z To Len(DaBa)
                If Mid(DaBa, A, 1) = "[" Then Exit For
                If Mid(DaBa, A, 1) = Chr(13) Or A = Len(DaBa) Then
                If A = Len(DaBa) Then Let A = A + 1
                CurData = Mid(DaBa, Z + 1, A - (Z + 1))
                'List1.AddItem CurData
                Exit For
                End If
               Next A
            
             End If
             Next Z
        Exit For
        End If
        Next x
    End If
Next y
End Sub

Sub extractMenu()
On Error Resume Next
For i = 1 To List1.ListCount
Load mnuArray(i)
   mnuArray(i).Caption = List1.List(i - 1)
Next i
End Sub
