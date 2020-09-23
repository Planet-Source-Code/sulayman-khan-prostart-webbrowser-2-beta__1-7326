VERSION 5.00
Begin VB.Form frmCredits 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About - Prostart Web Browser"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4695
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Command1"
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
      Left            =   3240
      TabIndex        =   3
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton cmdCredits 
      Caption         =   "Command1"
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
      Left            =   3600
      TabIndex        =   2
      Top             =   3600
      Width           =   375
   End
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picScroll 
      BackColor       =   &H00000000&
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
      Height          =   2895
      Left            =   0
      Picture         =   "frmCredits.frx":0000
      ScaleHeight     =   2895
      ScaleWidth      =   5055
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.TextBox txtScroll 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   1695
         Left            =   1560
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   2040
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "CLOSE"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   2640
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    left = (Screen.Width - Width) \ 2
    top = (Screen.Height - Height) \ 2
'Variable Declarations
    Dim iFileNum As Integer
    Dim lLineCount As Long
    Dim lLineHeight As Long
    
    On Error GoTo ErrHandler 'Goto to ErrHandler if an error occurs
    
    If cmdCredits.Caption = "Hide Credits" Then
        picScroll.Visible = False
        tmrScroll.Enabled = False
        cmdCredits.Caption = "&Roll Credits"
    Else
        iFileNum = FreeFile
        'open file and read text from it
        Open App.Path & "\credits.txt" For Input As iFileNum
        txtScroll = Input(LOF(iFileNum), iFileNum)
        Close #iFileNum 'close file
        lLineCount = SendMessage(txtScroll.hwnd, EM_GETLINECOUNT, 0&, 0&)
        lLineHeight = TextHeight("TEST") 'Get the height of text in file
        txtScroll.Height = lLineHeight * lLineCount
        picScroll.left = 0
        picScroll.Visible = True
        tmrScroll.Enabled = True
        cmdCredits.Caption = "Hide Credits"
    End If
    Exit Sub

ErrHandler:
    txtScroll.Text = "File Not Found !!!" & vbNewLine & "The Credits.txt file is missing"
    Resume Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label1.ForeColor = vbWhite
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label1.ForeColor = &HC0FFFF
End Sub

Private Sub tmrScroll_Timer()
    'scroll txtScroll
    If txtScroll.top + txtScroll.Height < picScroll.top Then 'picScroll.Top
        txtScroll.top = picScroll.Height
    Else
        txtScroll.top = txtScroll.top - 25
    End If
End Sub

Private Sub txtScroll_GotFocus()
    cmdOk.SetFocus
    'Don't let the text box get focus, althought the text
    'box is locked it looks bad to see a cursor in the
    'text box as it scrolls up
    
    
End Sub

