VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmGetMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail Checker"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   Icon            =   "frmGetMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox tName 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox tServer 
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cCheck 
      Caption         =   "Check Mail"
      Default         =   -1  'True
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin MSWinsockLib.Winsock wPOP 
      Left            =   3960
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mail Check"
      Height          =   2415
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4575
      Begin VB.Label Label4 
         Caption         =   "Note: This mail check is not automatically done."
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Password:"
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Username:"
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Server:"
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmGetMail.frx":030A
         Top             =   240
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmGetMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bFirstTime As Boolean
Dim PrevCMD As String

Private Sub cCheck_Click()
    wPOP.Close
    bFirstTime = True
    wPOP.Connect tServer.Text, 110
End Sub



Private Sub Form_Load()
tServer.Text = frmOptions.Text3
tName.Text = frmOptions.Text4
tPassword.Text = frmOptions.Text5
End Sub

Private Sub wPOP_DataArrival(ByVal bytesTotal As Long)
    Dim X As String
    Dim sTemp As String
    Dim iMessages As Integer
    
    On Error Resume Next
    
    'Get the response:
    wPOP.GetData X
    'Check to see if the response is positive:
    If left(X, 1) = "+" Then
        'Is this the first response?
        If bFirstTime Then
            'Yes it is, switch the flag:
            bFirstTime = False
            'Send the username:
            sTemp = "USER " & tName.Text & vbCrLf
            wPOP.SendData sTemp
            'Needed se we can know where we are:
            PrevCMD = "USER"
        
        'Not the first response from server:
        Else
            'What was the previous command we sent?
            Select Case PrevCMD
                Case "USER"
                    'We sent the username, now send the password:
                    sTemp = "PASS " & tPassword.Text & vbCrLf
                    wPOP.SendData sTemp
                    PrevCMD = "PASS"
                Case "PASS"
                    'We sent the password, get the number of messages:
                    X = UCase(X)
                    'Filter out the number of messages, the response is something like:
                    '   +OK username has 4 message(s) (640 octets)
                    
                    iMessages = Val(Mid$(X, InStr(X, " HAS ") + _
                        Len(" HAS "), InStr(X, "MESSAGE") - (InStr(X, " HAS ") _
                        + Len(" HAS "))))
                        
                    'We did what we had to do, close the connection:
                    wPOP.SendData "QUIT" & vbCrLf
                    wPOP.Close
                    
                    'Show the user how many messages we have
                    MsgBox "You have " & iMessages & " messages in your mailbox"
            End Select
        End If
    Else
        'We got a negative reply, show it to the user
        
        'Start from the fifth because we don't need to show
        'the -ERR and the space following it:
        MsgBox Mid(X, 5)
    End If
End Sub

