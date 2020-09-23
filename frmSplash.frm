VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2655
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   4695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   4560
      Top             =   2640
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Sulayman Khan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "BETA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2940
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 2.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   2655
      Left            =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   5205
      Left            =   0
      Picture         =   "frmSplash.frx":030A
      Top             =   0
      Width           =   1545
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PROSTART WEB BROWSER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1580
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ApplicationForm As Form
'*   NOTE : Set FrmSplash as the startup object form.    *
Private Sub Form_Load()
'*   NOTE : put the name of your application form below  *
'*          instead of 'FrmApp'                          *
    Set ApplicationForm = frmMain
    
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags
    
    Timer1.Interval = 1000 ' set timer interval to 1 second.

End Sub

Private Sub Form_Click()
closeSplash
End Sub


Private Sub Form_Unload(Cancel As Integer)
Timer1.Enabled = False 'Make sure the timer control is disabled
End Sub

Private Sub Image1_Click()
closeSplash
End Sub

Private Sub Label1_Click()
closeSplash
End Sub

Private Sub Label2_Click()
closeSplash
End Sub

Private Sub Label3_Click()
closeSplash
End Sub

Private Sub Timer1_Timer()
    ' make it all happen.
    Static TimerCount As Integer
    TimerCount = TimerCount + 1
    Select Case TimerCount
        Case 1: ' 1 second elapsed, load and show your application.
                Load ApplicationForm
                ApplicationForm.Visible = True
                ApplicationForm.Enabled = False
                Timer1.Interval = 1500 ' reset timer interval to 3 seconds.
        Case 2: ' 4 seconds elapsed, disable timer and unload splash screen.
                Timer1.Enabled = False
                Unload frmSplash
                ApplicationForm.Enabled = True
    End Select
End Sub


Sub closeSplash()                           'I put this in a sub for itself because it needs to be executed
    If ApplicationForm.Visible = True Then  'when clicked ANYWHERE on the splashScreen, not only on the form.
        Timer1.Enabled = False
        Unload frmSplash
        ApplicationForm.Enabled = True
    End If
End Sub
