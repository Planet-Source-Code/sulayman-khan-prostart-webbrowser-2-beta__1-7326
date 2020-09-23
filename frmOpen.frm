VERSION 5.00
Begin VB.Form frmOpen 
   Caption         =   "Open Page"
   ClientHeight    =   330
   ClientLeft      =   9045
   ClientTop       =   2655
   ClientWidth     =   2910
   Icon            =   "frmOpen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   330
   ScaleWidth      =   2910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdGO 
      Appearance      =   0  'Flat
      Caption         =   "GO"
      Default         =   -1  'True
      Height          =   310
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtURL 
      Height          =   310
      Left            =   0
      TabIndex        =   0
      Text            =   "http://"
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGO_Click()
On Error Resume Next
frmMain.WebBrowser1.Navigate Text1.Text
Me.Hide
End Sub

Private Sub Form_Load()
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, flags
txtURL.Text = frmMain.WebBrowser1.LocationURL
End Sub

Private Sub Form_Resize()
txtURL.Width = Me.ScaleWidth - cmdGO.Width - 50
cmdGO.left = Me.Width - cmdGO.Width - 110
If Me.Height <> 735 Then Me.Height = 735
If Me.Width < 3030 Then Me.Width = 3030
End Sub
