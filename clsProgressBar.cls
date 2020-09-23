VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *****************   clsProgress Bar   *******************
'
' Code by Sergei Lossev
' slossev@uwaterlo.ca
'
' I don't care how you decide to use this code as long as you use
' it for a good cause. :o)
'
'
' PS This code may not be fully optimized but I never had to time to make it optimized.
' This was simply a 30 min task.

Option Explicit

Public Enum PROGRESS_BAR_STYLES
    PB_SMOOTH = 0
    PB_PARTITIONED = 2
End Enum

Public Enum PROGRESS_BAR_BORDER_STYLES
    PBB_NONE = 0
    PBB_FLAT = 1
    PBB_INSET = 2
    PBB_RAISED = 3
End Enum

Private m_iSpaceWidth As Integer
Private m_iPartitionNumber As Integer

Private m_iPartitionWidth As Integer

Private m_iTX As Integer
Private m_iTY As Integer

Private m_eStyle As PROGRESS_BAR_STYLES
Private m_picContainer As PictureBox

Private m_lMin As Long
Private m_lMax As Long
Private m_lValue As Long
Private m_lSize As Long

Private m_lBackColor As Long
Private m_lFillColor As Long
Private m_lTextColor As Long

Private m_eBorderStyle As PROGRESS_BAR_BORDER_STYLES

Private m_iFontSize As Integer

Private m_bShowPercentage As Boolean

Public Property Get Min() As Long
    Min = m_lMin
End Property
Public Property Let Min(ByVal lNewValue As Long)
    If lNewValue >= m_lMax Then Exit Property
    
    m_lMin = lNewValue
    m_lSize = m_lMax - m_lMin
    DrawProgressBar
End Property

Public Property Get ShowPercentage() As Boolean
    ShowPercentage = m_bShowPercentage
End Property
Public Property Let ShowPercentage(ByVal bNewValue As Boolean)
    m_bShowPercentage = bNewValue
    
End Property

Public Property Get Max() As Long
    Max = m_lMax
End Property
Public Property Let Max(ByVal lNewValue As Long)
    If lNewValue <= m_lMin Then Exit Property
    m_lMax = lNewValue
    
    m_lSize = m_lMax - m_lMin
    DrawProgressBar
End Property

Public Property Get Value() As Long
    Value = m_lValue
End Property

Public Property Let Value(ByVal lNewValue As Long)
    If lNewValue > m_lMax Or lNewValue < m_lMin Then Exit Property
    'm_lValue = lNewValue
    m_lValue = lNewValue - m_lMin
    DrawProgressBar
End Property

Public Property Get Style() As PROGRESS_BAR_STYLES
    Style = m_eStyle
End Property

Public Property Let Style(ByVal eNewValue As PROGRESS_BAR_STYLES)
    m_eStyle = eNewValue
    m_picContainer.Line (m_iTX, m_iTY)-(m_picContainer.Width - m_iTX * 2, m_picContainer.Height - m_iTY * 2), m_lBackColor, BF
    If eNewValue = PB_SMOOTH Then m_bShowPercentage = True
    DrawProgressBar
End Property

Public Property Get BorderStyle() As PROGRESS_BAR_BORDER_STYLES
    BorderStyle = m_eBorderStyle
End Property

Public Property Let BorderStyle(ByVal eNewValue As PROGRESS_BAR_BORDER_STYLES)
    m_eBorderStyle = eNewValue
    DrawProgressBar
End Property

Public Property Get PictureBox() As Picture
    Set PictureBox = m_picContainer
End Property

Public Property Get FillColor() As Long
    FillColor = m_lFillColor
End Property
Public Property Let FillColor(ByVal lNewValue As Long)
    m_lFillColor = lNewValue
End Property

Public Property Get Backcolor() As Long
    Backcolor = m_lBackColor
End Property
Public Property Let Backcolor(ByVal lNewValue As Long)
    m_lBackColor = lNewValue
End Property

Public Property Get TextColor() As Long
    TextColor = m_lTextColor
End Property
Public Property Let TextColor(ByVal lNewValue As Long)
    m_lTextColor = lNewValue
End Property

Public Property Set PictureBox(picAny As PictureBox)
    Dim iCount As Integer
    Set m_picContainer = picAny
    
    With m_picContainer
        .BorderStyle = 0
        .AutoRedraw = True
        
        .DrawMode = 13
        .ScaleMode = vbTwips
        
        .Font = "Arial"
        .FontBold = True
        
        For iCount = 1 To 30
            .FontSize = iCount
            If .TextHeight("100%") > .Height * 3 / 4 Then
                Exit For
            End If
        Next iCount
        m_iFontSize = .FontSize
        
        m_iSpaceWidth = 3 * m_iTX
        m_iPartitionWidth = (.Width - m_iTX * 6 - (m_iPartitionNumber - 1) * m_iSpaceWidth) / 20
        
        
    End With
    
    
    DrawProgressBar
End Property

Public Sub DrawProgressBar()
    Dim iX1 As Integer
    Dim iX2 As Integer
    Dim iY1 As Integer
    Dim iY2 As Integer
    
    Dim iX As Integer
    
    Dim sPercent As String
    
    If m_picContainer Is Nothing Then
        MsgBox "Progress bar picture box is not binded!", vbOKOnly, "Progress Bar"
        Exit Sub
    End If
    
    With m_picContainer
    
        'Background
        m_picContainer.Line (CLng(.Width * m_lValue / (m_lSize) + m_iTX), m_iTY)-(.Width - m_iTX * 2, .Height - m_iTY * 2), m_lBackColor, BF
    
        Select Case m_eStyle
            Case PB_SMOOTH
                    
                    '* * * Progress * * *
                    iX1 = m_iTX * 2
                    iY1 = m_iTY * 2
                    iX2 = CLng((.Width - m_iTX * 4) * m_lValue / (m_lSize))
                    'iX1 = iX2
                    
                    iY2 = .Height - m_iTY * 3
                    If iX2 < iX1 Then iX2 = iX1
                    
                    If iX1 <> iX2 Then m_picContainer.Line (iX1, iY1)-(iX2, iY2), m_lFillColor, BF
                    
                    
            
            Case PB_PARTITIONED
                
                    '* * * Progress * * *
                    iX1 = m_iTX * 2
                    iY1 = m_iTY * 2
                    iX2 = CLng((.Width - m_iTX * 4) * m_lValue / (m_lSize))
                    iY2 = .Height - m_iTY * 3
                    
                    For iX = iX1 To iX2 Step m_iPartitionWidth + m_iSpaceWidth
                        
                        If iX2 < iX1 Then iX2 = iX1
                    
                        If iX1 <> iX2 Then
                            If iX + m_iPartitionWidth > iX2 Then
                                m_picContainer.Line (iX, iY1)-(iX2, iY2), m_lFillColor, BF
                            Else
                                m_picContainer.Line (iX, iY1)-(iX + m_iPartitionWidth, iY2), m_lFillColor, BF
                            End If
                        End If
                        
                    Next iX
                    
                    'm_picContainer.Line (iX - iX2 - m_iSpaceWidth - m_iPartitionWidth, iY1)-(iX2, iY2), m_lFillColor, BF
                    
                    m_bShowPercentage = False
                                     
        End Select
        
        'Draw the border
        Select Case m_eBorderStyle
            Case PBB_FLAT
                m_picContainer.Line (0, 0)-(.Width - m_iTX * 1, .Height - m_iTY * 1), vbBlack, B
            Case PBB_INSET
                m_picContainer.Line (0, 0)-(.Width - m_iTX * 1, 0), vb3DShadow
                m_picContainer.Line (0, 0)-(0, .Height - m_iTY), vb3DShadow
                m_picContainer.Line (.Width - m_iTX, 0)-(.Width - m_iTX, .Height - m_iTY), vb3DHighlight
                m_picContainer.Line (0, .Height - m_iTX)-(.Width, .Height - m_iTY), vb3DHighlight
                
            Case PBB_RAISED
                m_picContainer.Line (0, 0)-(.Width - m_iTX * 1, 0), vb3DHighlight
                m_picContainer.Line (0, 0)-(0, .Height - m_iTY * 1), vb3DHighlight
                m_picContainer.Line (.Width - m_iTX * 1, 0)-(.Width - m_iTX * 1, .Height - m_iTY), vb3DShadow
                m_picContainer.Line (0, .Height - m_iTY)-(.Width - m_iTX * 1, .Height - m_iTY), vb3DShadow
            Case Else
                m_picContainer.Line (0, 0)-(.Width - m_iTX * 1, .Height - m_iTY), m_lBackColor, B
        End Select
    
        
        ' Show the progress percentage
        If m_bShowPercentage Then
            .ForeColor = m_lTextColor
            .Font = "Arial"
            .FontBold = True
            
            .FontSize = m_iFontSize
            .DrawMode = vbCopyPen
            
            sPercent = CStr(Round(m_lValue / (m_lSize) * 100, 0) & "%")
            
            .CurrentX = (.Width - .TextWidth(sPercent)) / 2
            .CurrentY = (.Height - .TextHeight(sPercent)) / 2
    
            m_picContainer.Print sPercent
        
            
        End If
        
    End With
    
    DoEvents
    
End Sub

Private Sub Class_Initialize()
    m_eStyle = PB_SMOOTH
    
    m_eBorderStyle = PBB_INSET

    m_iPartitionNumber = 20
    
    m_lMax = 1
    m_lMin = 0
    m_lValue = 0
    m_lSize = 1
    m_lBackColor = vb3DFace
    m_lFillColor = vbHighlight
    
    m_iFontSize = 0
    
    m_bShowPercentage = True
    
    m_lTextColor = vbHighlightText
    
    m_iTX = Screen.TwipsPerPixelX
    m_iTY = Screen.TwipsPerPixelY
    
End Sub



