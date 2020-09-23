VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSource 
   Caption         =   "Document Source"
   ClientHeight    =   3870
   ClientLeft      =   3450
   ClientTop       =   2370
   ClientWidth     =   5070
   Icon            =   "frmSource.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   5070
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6800
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmSource.frx":030A
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
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu seperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    frmSource.Caption = frmMain.WebBrowser1.LocationURL
    left = (Screen.Width - Width) \ 2
    top = (Screen.Height - Height) \ 2
    Dim txt As String
    Dim b() As Byte
    
    b() = Inet1.OpenURL(frmMain.WebBrowser1.LocationURL, 1)
    
    txt = ""
    


    For t = 0 To UBound(b) - 1
        txt = txt + Chr(b(t))
    Next
    
    RichTextBox1.Text = txt

    
    Exit Sub

End Sub

Private Sub Form_Resize()
RichTextBox1.Height = frmSource.Height - 600
RichTextBox1.Width = frmSource.Width - 150

End Sub

Private Sub mnuClose_Click()
Unload Me
End Sub

Private Sub mnuCopy_Click()
    Clipboard.Clear 'delete everthing in the clipboard
    Clipboard.SetText RichTextBox1.SelText, 1 'put your text into it on place 1

End Sub

Private Sub mnuCut_Click()
    Clipboard.Clear 'delete everthing in the clipboard
    Clipboard.SetText RichTextBox1.SelText, 1 'put your text into it on place 1
    'If the place 1 allready excists it will be erased
    RichTextBox1.SelText = "" 'delete everyting that was selected in the textbox
End Sub

Private Sub mnuPaste_Click()
    RichTextBox1.SelText = Clipboard.GetText(1)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    PopupMenu mnuFile 'Your menu's name
End Sub

Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
    PopupMenu mnuFile 'Your menu's name
    End If
End Sub
