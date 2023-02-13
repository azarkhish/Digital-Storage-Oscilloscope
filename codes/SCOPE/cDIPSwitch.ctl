VERSION 5.00
Begin VB.UserControl cDIPSwitch 
   Appearance      =   0  'Flat
   CanGetFocus     =   0   'False
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2625
   FillColor       =   &H0000FF00&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   ScaleHeight     =   43
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   175
   ToolboxBitmap   =   "cDIPSwitch.ctx":0000
End
Attribute VB_Name = "cDIPSwitch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

Private Type eSwitch
    Guide       As RECT
    State       As Boolean
End Type
Private m_Switch() As eSwitch 'Switches are stored here top to bottom or left to right, regardless of m_ONOrientation

Public Event Click(SwitchID As Integer, NewState As Boolean)
Public Event MouseMove(SwitchID As Integer)

Public Enum eDIPAppearance
    daFlat = 0
    da3DRaised = 1
    da3DSunken = 2
    da3DEtched = 3
End Enum

Private m_Appearance        As eDIPAppearance
Private m_BackColor         As OLE_COLOR
Private m_FontChanged       As Boolean
Private m_GuideColor        As OLE_COLOR
Private m_ONOrientation     As AlignConstants
Private m_ShowON            As Boolean
Private m_SwitchOFFColor    As OLE_COLOR
Private m_SwitchONColor     As OLE_COLOR
Private m_SwitchCount       As Integer
Private m_SwitchGap         As Integer

Private Const m_def_Appearance = da3DRaised
Private Const m_def_Enabled = True
Private Const m_def_GuideColor = vbButtonShadow
Private Const m_def_ONOrientation = vbAlignTop
Private Const m_def_State = "00000000"
Private Const m_def_ShowON = True
Private Const m_def_SwitchOFFColor = vbYellow
Private Const m_def_SwitchONColor = vbYellow
Private Const m_def_SwitchCount = 8
Private Const m_def_SwitchGap = 5 'percent

Private Const s_Appearance = "Appearance"
Private Const s_BackColor = "BackColor"
Private Const s_Enabled = "Enabled"
Private Const s_Font = "Font"
Private Const s_ForeColor = "ForeColor"
Private Const s_GuideColor = "GuideColor"
Private Const s_ONOrientation = "ONOrientation"
Private Const s_State = "State"
Private Const s_ShowON = "ShowON"
Private Const s_SwitchOFFColor = "SwitchOFFColor"
Private Const s_SwitchONColor = "SwitchONColor"
Private Const s_SwitchCount = "SwitchCount"
Private Const s_SwitchGap = "SwitchGap"

Private Enum eDrawEdgeStyle
    BDR_RAISEDOUTER = &H1
    BDR_SUNKENOUTER = &H2
    BDR_RAISEDINNER = &H4
    BDR_SUNKENINNER = &H8
    BDR_RAISED = &H5
    BDR_SUNKEN = &HA
    BDR_OUTER = &H3
    BDR_INNER = &HC
    EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
    EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
    EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
    EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
    NO_EDGE = -1
End Enum

Private Enum eDrawEdgeFlags
    BF_LEFT = &H1
    BF_TOP = &H2
    BF_RIGHT = &H4
    BF_BOTTOM = &H8
    BF_MIDDLE = &H800       'Fill in the middle
    BF_SOFT = &H1000        'For softer buttons
    BF_ADJUST = &H2000      'Calculate the space left over
    BF_FLAT = &H4000        'For flat rather than 3D borders
    BF_MONO = &H8000        'For monochrome borders
    BF_RECT = BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM
End Enum

Private Enum eDrawTextFormat
    DT_LEFT = &H0
    DT_CENTER = &H1
    DT_RIGHT = &H2
    DT_WORDBREAK = &H10
    DT_CALCRECT = &H400
    DT_SINGLELINE = &H20
    DT_TOP = &H0 Or DT_SINGLELINE
    DT_VCENTER = &H4 Or DT_SINGLELINE
    DT_BOTTOM = &H8 Or DT_SINGLELINE
    DI_MASK = &H1
    DI_IMAGE = &H2
    DI_NORMAL = &H3
    DI_COMPAT = &H4
    DI_DEFAULTSIZE = &H8
End Enum

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hbrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long


Private Sub DrawSwitch(ByRef index As Integer)
    Dim swRECT      As RECT
    Dim txtRECT     As RECT
    Dim txtht       As Long
    Dim r           As Long
    Dim offset3d    As Integer
    Dim wd          As Single   'Single prevents rounding errors since everything's in pixels
    Dim ht          As Single   'Single prevents rounding errors since everything's in pixels
    Dim hbrush      As Long
    Dim color       As Long     'Colors must be translated from OLE_COLOR since APIs do not recognize VB system colors
    Dim ONwd        As Long     'defaults to zero
    Dim ONht        As Long     'defaults to zero
    
    offset3d = IIf(m_Appearance <> daFlat, 4, 0)
    If (m_ShowON = True) Then
        ONwd = UserControl.TextWidth("ON") + 4
        ONht = UserControl.TextHeight("ON") + 4
    End If
    
    With m_Switch(index).Guide
        Select Case m_ONOrientation
            Case vbAlignTop, vbAlignBottom       'Horizontal
                wd = (UserControl.ScaleWidth - IIf(m_Appearance = daFlat, 0, 2 * offset3d) - ONwd) / (m_SwitchCount + ((m_SwitchCount - 1) * (m_SwitchGap / 100)))
                .Left = offset3d + (index - 1) * (wd * (1 + (m_SwitchGap / 100)))
                .Top = offset3d
                .Right = .Left + CLng(wd)
                .Bottom = UserControl.ScaleHeight - offset3d
            
            Case vbAlignLeft, vbAlignRight    'Vertical
                ht = (UserControl.ScaleHeight - IIf(m_Appearance = daFlat, 0, 2 * offset3d) - ONht) / (m_SwitchCount + ((m_SwitchCount - 1) * (m_SwitchGap / 100)))
                .Left = offset3d
                .Top = offset3d + (index - 1) * (ht * (1 + (m_SwitchGap / 100)))
                .Bottom = .Top + CLng(ht)
                .Right = UserControl.ScaleWidth - offset3d
        End Select
    End With
    
    If ((index >= LBound(m_Switch)) And (index <= UBound(m_Switch))) Then
        With m_Switch(index)
            'Draw the guide
            color = TranslateColor(m_GuideColor)
            hbrush = CreateSolidBrush(color)
            r = FillRect(UserControl.hdc, .Guide, hbrush)
            r = DrawEdge(UserControl.hdc, .Guide, BDR_SUNKEN, BF_RECT)
            r = DeleteObject(hbrush)

            'Draw the switch
            swRECT = .Guide
            r = InflateRect(swRECT, -2, -2)
            
            Select Case m_ONOrientation
                Case vbAlignTop, vbAlignBottom       'Horizontal
                    swRECT.Bottom = swRECT.Top + (swRECT.Bottom - swRECT.Top) / 2
                    r = OffsetRect(swRECT, 0, IIf(.State = IIf(m_ONOrientation = vbAlignTop, False, True), .Guide.Bottom - swRECT.Bottom - 2, 0))
                Case vbAlignLeft, vbAlignRight    'Vertical
                    swRECT.Right = swRECT.Left + (swRECT.Right - swRECT.Left) / 2
                    r = OffsetRect(swRECT, IIf(.State = IIf(m_ONOrientation = vbAlignLeft, False, True), .Guide.Right - swRECT.Right - 2, 0), 0)
            End Select
            
            color = TranslateColor(IIf(.State = True, m_SwitchONColor, m_SwitchOFFColor))
            hbrush = CreateSolidBrush(color)
            r = FillRect(UserControl.hdc, swRECT, hbrush)
            r = DrawEdge(UserControl.hdc, swRECT, BDR_RAISED, BF_RECT)
            r = DeleteObject(hbrush)

            'Draw the text
            txtRECT = swRECT
            r = DrawText(UserControl.hdc, CStr(SwitchNumber(index)), -1, txtRECT, DT_CENTER Or DT_VCENTER)
        End With
    End If
End Sub
Private Function FindSwitch(X As Single, Y As Single) As Integer
    Dim i As Integer
    
    On Error GoTo notswitch
    
    FindSwitch = -1
    
    'You could use the PtInRect API function, but I can not get it to work.
    For i = LBound(m_Switch) To UBound(m_Switch) 'Generates an error if not dimensioned
        With m_Switch(i).Guide
            If ((X >= .Left) And (X < .Right) And _
                (Y >= .Top) And (Y < .Bottom)) Then
                FindSwitch = i
                Exit For
            End If
        End With
    Next i
notswitch:
End Function
Public Sub SetState(firstswitch As Integer, states As String)
    Dim i As Integer
    
    On Error GoTo byebye
    
    For i = 1 To Len(states)
        m_Switch(SwitchNumber(firstswitch + i - 1)).State = IIf(Mid$(states, i, 1) = "1", True, False)
    Next i
byebye:
    UserControl.Refresh
End Sub
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByRef NewFont As Font)
    Set UserControl.Font = NewFont
    m_FontChanged = True
    UserControl.Refresh
    PropertyChanged s_Font
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property
Public Property Get SwitchONColor() As OLE_COLOR
    SwitchONColor = m_SwitchONColor
End Property
Public Property Get SwitchOFFColor() As OLE_COLOR
    SwitchOFFColor = m_SwitchOFFColor
End Property

'
Public Property Get GuideColor() As OLE_COLOR
    GuideColor = m_GuideColor
End Property
Public Property Get SwitchCount() As Integer
    SwitchCount = m_SwitchCount
End Property
Public Property Get SwitchGap() As String
Attribute SwitchGap.VB_Description = "Width of gap between switches (as a percentage of switch width).  100 means gap is 100%, (ie, the same size ) of the width of the switch itself."
    SwitchGap = Format(m_SwitchGap / 100, "0%")
End Property
Public Property Let BackColor(newcolor As OLE_COLOR)
    m_BackColor = newcolor
    UserControl.BackColor = m_BackColor
    UserControl.Refresh
    PropertyChanged s_BackColor
End Property
Public Property Let SwitchONColor(newcolor As OLE_COLOR)
    m_SwitchONColor = newcolor
    UserControl.Refresh
    PropertyChanged s_SwitchONColor
End Property
Public Property Let SwitchOFFColor(newcolor As OLE_COLOR)
    m_SwitchOFFColor = newcolor
    UserControl.Refresh
    PropertyChanged s_SwitchOFFColor
End Property

Public Property Let SwitchGap(newgap As String)
    m_SwitchGap = IIf(CSng(newgap) < 0, 0, CSng(newgap))
    UserControl.Refresh
    PropertyChanged s_SwitchGap
End Property
Public Property Let ForeColor(newcolor As OLE_COLOR)
    UserControl.ForeColor = newcolor
    UserControl.Refresh
    PropertyChanged s_ForeColor
End Property
Public Property Let Enabled(isenabled As Boolean)
    UserControl.Enabled = isenabled
    PropertyChanged s_Enabled
End Property
Public Property Let GuideColor(newcolor As OLE_COLOR)
    m_GuideColor = newcolor
    UserControl.Refresh
    PropertyChanged s_GuideColor
End Property
Public Property Let SwitchCount(new_count As Integer)
    If ((new_count >= 1) And (new_count <> m_SwitchCount)) Then
        ReDim m_Switch(1 To new_count)
        m_SwitchCount = new_count
    End If
    
    UserControl.Refresh
    PropertyChanged s_SwitchCount
End Property
Public Property Get Appearance() As eDIPAppearance
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Appearance.VB_UserMemId = -520
    Appearance = m_Appearance
End Property
Public Property Get ShowON() As Boolean
    ShowON = m_ShowON
End Property

Public Property Get ONOrientation() As AlignConstants
    ONOrientation = m_ONOrientation
End Property

Public Property Get State() As String
    Dim i   As Integer
    Dim txt As String
    
    For i = LBound(m_Switch) To UBound(m_Switch)
        txt = txt & IIf(m_Switch(i).State = True, "1", "0")
    Next i
    State = IIf((m_ONOrientation = vbAlignTop) Or (m_ONOrientation = vbAlignRight), txt, StrReverse(txt))
End Property
Public Property Let Appearance(newval As eDIPAppearance)
    m_Appearance = newval
    UserControl.Refresh
    PropertyChanged s_Appearance
End Property
Public Property Let ShowON(showit As Boolean)
    m_ShowON = showit
    UserControl.Refresh
    PropertyChanged s_ShowON
End Property
Public Property Let ONOrientation(orient As AlignConstants)
    Dim cur_vert    As Boolean
    Dim new_vert    As Boolean
    
    orient = IIf(orient = vbAlignNone, vbAlignTop, orient)
    
    cur_vert = IIf((m_ONOrientation = vbAlignLeft) Or (m_ONOrientation = vbAlignRight), True, False)
    new_vert = IIf((orient = vbAlignLeft) Or (orient = vbAlignRight), True, False)
    
    m_ONOrientation = orient
    If (cur_vert <> new_vert) Then 'ONOrientation changed from horizontal to vertical (or visa versa)
        Call UserControl.Size(UserControl.Height, UserControl.Width)
    End If
    
    UserControl.Refresh
    PropertyChanged s_ONOrientation
End Property
Public Property Let State(states As String)
    Dim i As Integer
    
    For i = LBound(m_Switch) To UBound(m_Switch)
        m_Switch(SwitchNumber(i)).State = IIf(Mid$(states, i - LBound(m_Switch) + 1, 1) = "1", True, False)
    Next i
    
    UserControl.Refresh
    PropertyChanged s_State
End Property
Private Function SwitchNumber(ByVal index As Integer) As Integer
    If ((m_ONOrientation = vbAlignBottom) Or (m_ONOrientation = vbAlignLeft)) Then
        SwitchNumber = UBound(m_Switch) + LBound(m_Switch) - index
    Else
        SwitchNumber = index
    End If
End Function
Private Function TranslateColor(color As OLE_COLOR) As Long
    'Using OLE_COLOR as the type for colors allows the user to choose normal colors
    '*and* SystemColorConstants (like vbButtonFace).  API drawing functions do not
    'recognize these system colors, so they must be translated to "real" colors.
    'All SystemColorConstants range in values from &H80000000 to &H80000018
    'GetSysColor only wants the lower byte of data (values 0-24), so you have to
    'mask off the upper part.  This makes SystemColorConstants equivalent to the
    'API-declared COLOR_xxxx constants
    If (color <= &H80000018) Then 'Could also test for (color >= &H80000000), but since this is the largest negative long possible, it would always be true
        TranslateColor = GetSysColor(color And &H1F)
    Else
        TranslateColor = color
    End If
End Function
Private Sub UserControl_Initialize()
    UserControl.ScaleMode = vbPixels
End Sub
Private Sub UserControl_InitProperties()
    Appearance = m_def_Appearance
    BackColor = Ambient.BackColor
    Enabled = m_def_Enabled
    Font = Ambient.Font
    ForeColor = Ambient.ForeColor
    GuideColor = m_def_GuideColor
    ONOrientation = m_def_ONOrientation
    ShowON = m_def_ShowON
    SwitchOFFColor = m_def_SwitchOFFColor
    SwitchONColor = m_def_SwitchONColor
    SwitchCount = m_def_SwitchCount
    SwitchGap = m_def_SwitchGap
    State = m_def_State
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sw As Integer
    
    sw = FindSwitch(X, Y)
    If ((sw >= LBound(m_Switch)) And (sw <= UBound(m_Switch))) Then
        RaiseEvent MouseMove(SwitchNumber(sw))
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sw As Integer
    
    sw = FindSwitch(X, Y)
    If ((sw >= LBound(m_Switch)) And (sw <= UBound(m_Switch))) Then
        m_Switch(sw).State = Not m_Switch(sw).State
        UserControl.Refresh
        RaiseEvent Click(SwitchNumber(sw), m_Switch(sw).State)
    End If
End Sub

Private Sub UserControl_Paint()
    Dim i           As Integer
    Dim clientRECT  As RECT
    Dim txtRECT     As RECT
    Dim ONwd        As Long     'defaults to zero
    Dim ONht        As Long     'defaults to zero
    Dim r           As Long
    Dim offset3d    As Long
    
    With UserControl
        r = SetRect(clientRECT, 0, 0, .ScaleWidth, .ScaleHeight)
        Select Case m_Appearance 'DrawEdge requires clientRECT in pixels
            Case daFlat:
            Case da3DRaised: DrawEdge .hdc, clientRECT, CLng(EDGE_RAISED), BF_RECT
            Case da3DSunken: DrawEdge .hdc, clientRECT, CLng(EDGE_SUNKEN), BF_RECT
            Case da3DEtched: DrawEdge .hdc, clientRECT, CLng(EDGE_ETCHED), BF_RECT
        End Select
        
        'Draw the ON label
        If (m_ShowON = True) Then
            offset3d = IIf(m_Appearance <> daFlat, 4, 0)
            ONwd = .TextWidth("ON") + 4
            ONht = .TextHeight("ON") + 4
            
            Select Case m_ONOrientation
                Case vbAlignTop:    r = SetRect(txtRECT, .ScaleWidth - ONwd - offset3d, 0, .ScaleWidth, .ScaleHeight / 2)
                Case vbAlignBottom:  r = SetRect(txtRECT, .ScaleWidth - ONwd - offset3d, .ScaleHeight / 2, .ScaleWidth, .ScaleHeight)
                Case vbAlignLeft:  r = SetRect(txtRECT, 0, .ScaleHeight - ONht - offset3d, .ScaleWidth / 2, .ScaleHeight)
                Case vbAlignRight: r = SetRect(txtRECT, .ScaleWidth / 2, .ScaleHeight - ONht - offset3d, .ScaleWidth, .ScaleHeight)
            End Select
            r = DrawText(.hdc, "ON", -1, txtRECT, DT_CENTER Or DT_VCENTER)
        End If
        
        For i = LBound(m_Switch) To UBound(m_Switch)
            DrawSwitch (i)
        Next i
    End With
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Appearance = .ReadProperty(s_Appearance, m_def_Appearance)
        BackColor = .ReadProperty(s_BackColor, Ambient.BackColor)
        Enabled = .ReadProperty(s_Enabled, m_def_Enabled)
        Set Font = .ReadProperty(s_Font, Ambient.Font)
        ForeColor = .ReadProperty(s_ForeColor, Ambient.ForeColor)
        GuideColor = .ReadProperty(s_GuideColor, m_def_GuideColor)
        ONOrientation = .ReadProperty(s_ONOrientation, m_def_ONOrientation)
        ShowON = .ReadProperty(s_ShowON, m_def_ShowON)
        SwitchOFFColor = .ReadProperty(s_SwitchOFFColor, m_def_SwitchOFFColor)
        SwitchONColor = .ReadProperty(s_SwitchONColor, m_def_SwitchONColor)
        SwitchCount = .ReadProperty(s_SwitchCount, m_def_SwitchCount)
        SwitchGap = .ReadProperty(s_SwitchGap, m_def_SwitchGap)
        State = .ReadProperty(s_State, m_def_State) 'Must be done last, since this is dependent on SwitchCount and ONOrientation
    End With
End Sub
Private Sub UserControl_Resize()
    Dim Xscale      As Single
    Dim Yscale      As Single
    Static prev_wd  As Single
    Static prev_ht  As Single
    Static fsize    As Single
    
    With UserControl
        If (fsize = 0) Then fsize = .Font.Size
        
        If ((.ScaleHeight <> prev_wd) Or (.ScaleWidth <> prev_ht) Or (prev_wd = prev_ht)) Then 'ONOrientation not changed
            If (m_FontChanged = False) Then 'Not just a manual Font change
                Xscale = .ScaleWidth / IIf(prev_wd <> 0, prev_wd, .ScaleWidth)    'Ratio current:previous scalewidth
                Yscale = .ScaleHeight / IIf(prev_ht <> 0, prev_ht, .ScaleHeight)  'Ratio current:previous scaleheight
                If (Xscale <> 1 And Yscale <> 1) Then 'If just changing one dimension, don't change size
                    fsize = fsize * IIf(Xscale < Yscale, Xscale, Yscale)
                    .FontSize = Round(fsize)
                End If
            End If
        End If
        m_FontChanged = False
        prev_wd = .ScaleWidth
        prev_ht = .ScaleHeight
    End With
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty(s_Appearance, m_Appearance, m_def_Appearance)
        Call .WriteProperty(s_BackColor, m_BackColor, Ambient.BackColor)
        Call .WriteProperty(s_Enabled, UserControl.Enabled, m_def_Enabled)
        Call .WriteProperty(s_Font, UserControl.Font, Ambient.Font)
        Call .WriteProperty(s_ForeColor, UserControl.ForeColor, Ambient.ForeColor)
        Call .WriteProperty(s_GuideColor, m_GuideColor, m_def_GuideColor)
        Call .WriteProperty(s_ONOrientation, m_ONOrientation, m_def_ONOrientation)
        Call .WriteProperty(s_ShowON, m_ShowON, m_def_ShowON)
        Call .WriteProperty(s_State, State, m_def_State)
        Call .WriteProperty(s_SwitchOFFColor, m_SwitchOFFColor, m_def_SwitchOFFColor)
        Call .WriteProperty(s_SwitchONColor, m_SwitchONColor, m_def_SwitchONColor)
        Call .WriteProperty(s_SwitchCount, m_SwitchCount, m_def_SwitchCount)
        Call .WriteProperty(s_SwitchGap, m_SwitchGap, m_def_SwitchGap)
    End With
End Sub
