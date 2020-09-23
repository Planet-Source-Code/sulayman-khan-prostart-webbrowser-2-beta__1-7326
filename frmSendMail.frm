VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSendMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send E-mail"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSendMail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3000
      TabIndex        =   11
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Default         =   -1  'True
      Height          =   495
      Left            =   1920
      TabIndex        =   10
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtMessage 
      Height          =   2085
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1920
      Width           =   4035
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   1200
      Width           =   2355
   End
   Begin VB.TextBox txtToEmail 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   840
      Width           =   2355
   End
   Begin VB.TextBox txtFromEmail 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   2355
   End
   Begin VB.TextBox txtServerDomain 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   2355
   End
   Begin MSWinsockLib.Winsock w 
      Left            =   7560
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label7 
      Caption         =   "Message:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Subject:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "To Email:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "From Email:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Server Domain:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'SMTP Mailer
'By Ara Mahdessian (ara@dev-center.com)
'
'go to www.dev-center.com for VB "stuff"

Private Response As String

Sub SendEmail(ServerDomain As String, FromEmail As String, ToEmail As String, Subject As String, Body As String)
    
    w.LocalPort = 0 ' Must set local port to 0 (Zero) or you can only send 1 e-mail per program start
    
    If w.State <> sckClosed Then w.Close 'close winsock if open
    
    w.Protocol = sckTCPProtocol 'use tcp/ip protocol
    w.RemoteHost = ServerDomain 'server domain
    w.RemotePort = 25           '25 is standard smtp port
    w.Connect                   'connect
    
    WaitForResponse ("220")             'wait for confirmed connection
    
    w.SendData "HELO " & ServerDomain & vbCrLf   'send HELO msg
    WaitForResponse ("250")                             'wait for response
    
    w.SendData "MAIL FROM: <" & FromEmail & ">" & vbCrLf  'sender's email
    WaitForResponse ("250")                                         'wait for response
    
    w.SendData "RCPT TO: <" & ToEmail & ">" & vbCrLf   'recipient's email
    WaitForResponse ("250")                                     'wait for response
    
    w.SendData ("data" & vbCrLf)    'tell server msg and headers are incoming
    
    WaitForResponse ("354")
    w.SendData "From: " & FromEmail & vbCrLf        'name of sender
    w.SendData "X-Mailer: Dev-Center SMTP Mailer" & vbCrLf 'name of program [customize it]
    w.SendData "To: " & ToEmail & vbCrLf 'name of recipient
    w.SendData "Subject: " & Subject & vbCrLf 'subject of email
    
    w.SendData Body & vbCrLf    'send body (message)
    
    w.SendData "." & vbCrLf     'terminate incoming data/headers
    WaitForResponse ("250")     'wait for sent mail confirmation
    
    w.SendData "quit" & vbCrLf  'say bye-bye (quit)
    WaitForResponse ("221")     'wait for server to log you off - ethics folks
    
    w.Close
    
    MsgBox "Mail sent!", vbExclamation, "Mail Sent successfully"
    
End Sub


Sub WaitForResponse(ResponseCode As String)
    
    Dim Reply As Integer
    Dim Start As Single
    Dim Tmr As Single

    Start = Timer 'time in case server doesn't respond

    While Len(Response) = 0 'do until we get a response from server
        Tmr = Start - Timer 'get elapsed time

        DoEvents 'let system check for incoming response

        If Tmr > 10 Then 'if server is not responding (timed out)
            MsgBox "Error:" + vbCrLf + "SMTP service timed out while waiting for response!", vbExclamation, "SMTP Service Error"
            Exit Sub
        End If
    Wend


    While left(Response, 3) <> ResponseCode
        DoEvents
        
        If Tmr > 10 Then
            MsgBox "Error:" + vbCrLf + "Improper response code received: " + Response + vbCrLf + "Expected code: " + ResponseCode, vbExclamation, "SMTP Service Error"
            Exit Sub
        End If
    Wend
    
    Response = "" 'set response code to blank
    
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSend_Click()

    SendEmail txtServerDomain, txtFromEmail, txtToEmail, txtSubject, txtMessage

End Sub

Private Sub Form_Load()
Me.txtFromEmail = frmOptions.Text2
End Sub

Private Sub w_DataArrival(ByVal bytesTotal As Long)

    w.GetData Response 'get incoming response

End Sub

