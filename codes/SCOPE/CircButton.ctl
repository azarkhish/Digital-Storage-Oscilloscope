VERSION 5.00
Begin VB.UserControl CircButton 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   735
   DrawStyle       =   6  'Inside Solid
   DrawWidth       =   4
   PropertyPages   =   "CircButton.ctx":0000
   ScaleHeight     =   780
   ScaleWidth      =   735
End
Attribute VB_Name = "CircButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'API Declares for setting the button shape as round
'these functions cut the usercontrol and give it a round shape
'so that only the circle is opaque, the rest of the region is transparent

Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private cFace As Long, cLight As Long, cHighLight As Long, cShadow As Long, cDarkShadow As Long, cText As Long

Private Const PR_COLOR_BTNFACE = 15
Private Const PR_COLOR_BTNSHADOW = 16
Private Const PR_COLOR_BTNTEXT = 18
Private Const PR_COLOR_BTNHIGHLIGHT = 20
Private Const PR_COLOR_BTNDKSHADOW = 21
Private Const PR_COLOR_BTNLIGHT = 22

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
'Default Property Values:
Const m_def_ForeColor = 0
Const m_def_Caption = "CircBtn"
Const m_def_Radius = 0
'Property Variables:
Dim m_ForeColor As OLE_COLOR
Dim m_Caption As String
Dim m_Radius As Integer
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp

Private Sub SetColors()
'this function is taken form another upload in PSC,
'the Gurhan Button by 'Gurhan KARTAL
'Thanks Gurhan

    cFace = GetSysColor(PR_COLOR_BTNFACE)
    cShadow = GetSysColor(PR_COLOR_BTNSHADOW)
    cLight = GetSysColor(PR_COLOR_BTNLIGHT)
    cDarkShadow = GetSysColor(PR_COLOR_BTNDKSHADOW)
    cHighLight = GetSysColor(PR_COLOR_BTNHIGHLIGHT)
    cText = GetSysColor(PR_COLOR_BTNTEXT)

End Sub

Private Sub UserControl_Click()

    RaiseEvent Click
    draw "FocusRect"

End Sub

Private Sub UserControl_GotFocus()

    draw "FocusRect"

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyDown(KeyCode, Shift)
    Select Case KeyCode
      Case 32
        Call UserControl_MouseDown(0, 0, 0, 0)
      Case Else
    End Select

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyUp(KeyCode, Shift)
    Select Case KeyCode
      Case 32
        Call UserControl_MouseUp(0, 0, 0, 0)
        Call UserControl_Click
      Case Else
    End Select

End Sub

Private Sub UserControl_LostFocus()

    Cls
    UserControl_Resize

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseDown(Button, Shift, X, Y)
    draw "MouseDown"

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseUp(Button, Shift, X, Y)
    draw "Ordinary"

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000000)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_Radius = PropBag.ReadProperty("Radius", m_def_Radius)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    SetColors
    draw "Ordinary"
End Sub

Private Function draw(strArg As String, Optional bfocusRectMDwn As Boolean = False)
'this function draws the button. it uses the circle method to draw
'========================================================================================
'written by Praveen Menon
'12th Mrach 2002
'Last Modified Date
'12th March 2002

'========================================================================================
'arguments
'========================================================================================
    'strArg ==> used to determine which state of the button is to be drawn
    'three arguments can be there
        '   1) FocusRect ===> draws the focus rect of the button
        '   2) Ordinary ===> draws the ordinary or the mouseup state of the button
        '   3) MouseDown ===> draws the mousedown state or the keydown state of thebutton
    'bfocusRectMDwn ===> used to determine whether the focus rect is for
        '                the mousedown state or the mouseup state
'========================================================================================

  Dim rad1 As Integer
  Dim rad2 As Integer

    Select Case strArg
    
      Case "FocusRect"
      
      If Not bfocusRectMDwn Then
        UserControl.DrawStyle = 2
        DrawWidth = 1
        rad1 = Width / 2 - 100
        rad2 = Width / 2.5
        If rad1 < rad2 Then
            UserControl.Circle (Width / 2 - 30, Height / 2), (rad2), vbBlack
          Else 'NOT RAD1...
            UserControl.Circle (Width / 2 - 30, Height / 2), (rad1), vbBlack
        End If
        UserControl.DrawStyle = 6
        DrawWidth = 4
     Else
        UserControl.DrawStyle = 2
        DrawWidth = 1
        rad1 = Width / 2 - 100
        rad2 = Width / 2.5
        If rad1 < rad2 Then
            UserControl.Circle (Width / 2 - 5, Height / 2), (rad2), vbBlack
          Else 'NOT RAD1...
            UserControl.Circle (Width / 2 - 5, Height / 2), (rad1), vbBlack
        End If
        UserControl.DrawStyle = 6
        DrawWidth = 4
     End If
     
      Case "Ordinary"
        Cls
        UserControl.Circle (Width / 2 + 1, Height / 2 + 7), Width / 2, vbWhite
        UserControl.Circle (Width / 2 - 50, Height / 2), Width / 2, cDarkShadow
        UserControl.Circle (Width / 2 - 20, Height / 2), Width / 2, cShadow
        UserControl.ForeColor = m_ForeColor
        TextOut UserControl.hdc, (ScaleWidth - UserControl.TextWidth(m_Caption)) / 2 / Screen.TwipsPerPixelX, (ScaleHeight - UserControl.TextHeight(m_Caption)) / 2 / Screen.TwipsPerPixelY, m_Caption, Len(m_Caption)
      
      Case "MouseDown"
        Cls
        UserControl.Circle (Width / 2 + 1, Height / 2 + 10), Width / 2, cShadow
        DrawWidth = 4
        UserControl.ForeColor = m_ForeColor
        TextOut UserControl.hdc, (ScaleWidth - UserControl.TextWidth(m_Caption)) / 2 / Screen.TwipsPerPixelX + 2, (ScaleHeight - UserControl.TextHeight(m_Caption)) / 2 / Screen.TwipsPerPixelY + 2, m_Caption, Len(m_Caption)
        draw "FocusRect", True
      
      Case Else

    End Select

End Function

Private Sub UserControl_Resize()
Dim lrgn As Long

    Cls
    Width = Height
  
    lrgn = CreateEllipticRgn(0, 0, UserControl.Width / Screen.TwipsPerPixelX, UserControl.Height / Screen.TwipsPerPixelY)
    SetWindowRgn UserControl.hWnd, lrgn, True
    draw "Ordinary"
    m_Radius = Height

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."

    BackColor = UserControl.BackColor

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)

    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    draw "Ordinary"

End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "General"

    Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512

    Set Font = UserControl.Font

End Property

Public Property Set Font(ByVal New_Font As Font)

    Set UserControl.Font = New_Font
    PropertyChanged "Font"

End Property

Private Sub UserControl_KeyPress(KeyAscii As Integer)

    RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."

    Set MouseIcon = UserControl.MouseIcon

End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)

    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."

    MousePointer = UserControl.MousePointer

End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)

    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = "General"

    Caption = m_Caption

End Property

Public Property Let Caption(ByVal New_Caption As String)

    m_Caption = New_Caption
    PropertyChanged "Caption"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Radius() As Integer
Attribute Radius.VB_ProcData.VB_Invoke_Property = "General"

    Radius = m_Radius

End Property

Public Property Let Radius(ByVal New_Radius As Integer)

    m_Radius = New_Radius
    PropertyChanged "Radius"
    UserControl.Height = New_Radius
    'draw

End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

    Set UserControl.Font = Ambient.Font
    m_Caption = m_def_Caption
    m_Radius = m_def_Radius

    m_ForeColor = m_def_ForeColor
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000000)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("Radius", m_Radius, m_def_Radius)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
'    UserControl.ForeColor = m_ForeColor
    draw "Ordinary"
End Property

