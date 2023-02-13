VERSION 5.00
Begin VB.UserControl Wheel 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1515
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   ScaleHeight     =   101
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   101
   ToolboxBitmap   =   "Wheel.ctx":0000
   Begin VB.PictureBox picV_MSK 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H80000008&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   180
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   435
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picH_MSK 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H80000008&
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   435
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox picV_SND 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   30
      Picture         =   "Wheel.ctx":0312
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   435
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox picWHEEL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   120
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   435
      Width           =   180
   End
   Begin VB.PictureBox picH_SND 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   435
      Picture         =   "Wheel.ctx":0808
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Timer tmrRepeat 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   480
   End
   Begin VB.Shape shpWheelColour 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   1530
      Left            =   0
      Top             =   0
      Width           =   1530
   End
   Begin VB.Image imgRIGHT 
      Height          =   420
      Left            =   1200
      Picture         =   "Wheel.ctx":0CEE
      Top             =   0
      Width           =   315
   End
   Begin VB.Image imgVBack 
      Height          =   885
      Left            =   0
      Picture         =   "Wheel.ctx":1430
      Top             =   315
      Width           =   420
   End
   Begin VB.Image imgUP 
      Height          =   315
      Left            =   0
      Picture         =   "Wheel.ctx":27CE
      Top             =   0
      Width           =   420
   End
   Begin VB.Image imgHBack 
      Height          =   420
      Left            =   315
      Picture         =   "Wheel.ctx":2EF4
      Top             =   0
      Width           =   885
   End
   Begin VB.Image imgLEFT 
      Height          =   420
      Left            =   0
      Picture         =   "Wheel.ctx":42E6
      Top             =   0
      Width           =   315
   End
   Begin VB.Image imgDOWN 
      Height          =   315
      Left            =   0
      Picture         =   "Wheel.ctx":4A28
      Top             =   1200
      Width           =   420
   End
End
Attribute VB_Name = "Wheel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

    Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
    
    Dim Increment As Integer
    Dim GRABBED As Boolean
    Dim LastX As Single
    Dim LastY As Single
    Dim WheelPosition As Integer

'Default Property Values:
    Const m_def_Max = 100
    Const m_def_Min = 0
    Const m_def_Orientation = "V"
    Const m_def_ShowButtons = True
    Const m_def_ShadeControl = &H80000005
    Const m_def_ShadeWheel = &H80000005
    Const m_def_SpinOver = True
    Const m_def_Value = 0
    
'Property Variables:
    Dim m_Max As Integer
    Dim m_Min As Integer
    Dim m_Orientation As String
    Dim m_ShowButtons As Boolean
    Dim m_ShadeControl As OLE_COLOR
    Dim m_ShadeWheel As OLE_COLOR
    Dim m_SpinOver As Boolean
    Dim m_Value As Integer

'Event Declarations:
    Event Change()
    Event SpinLess()
    Event SpinMore()

'Press the DOWN button
Private Sub imgDOWN_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgDOWN.Move 2, 82 'Display the DOWN button in the pressed position
    Increment = 3 'Indicates the WHEEL is to be rotated by 3 pixels
    Call Wheel_Move 'Move the image of the WHEEL
    tmrRepeat.Enabled = True 'Enables REPEAT function if key is held down
End Sub

'Release the DOWN button
Private Sub imgDOWN_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tmrRepeat.Enabled = False 'Disables the REPEAT function
    imgDOWN.Move 0, 80 'Display the DOWN button in the raised position
End Sub

'Press the UP button
Private Sub imgUP_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgUP.Move 2, 2 'Display the UP button in the pressed position
    Increment = -3 'Indicates the WHEEL is to be rotated by -3 pixels
    Call Wheel_Move 'Move the image of the WHEEL
    tmrRepeat.Enabled = True 'Enables REPEAT function if key is held down
End Sub

'Release the UP button
Private Sub imgUP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tmrRepeat.Enabled = False 'Disables the REPEAT function
    imgUP.Move 0, 0 'Display the UP button in the raised position
End Sub

'Press the RIGHT button
Private Sub imgRIGHT_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgRIGHT.Move 82, 2 'Display the RIGHT button in the pressed position
    Increment = 3 'Indicates the WHEEL is to be rotated by 3 pixels
    Call Wheel_Move 'Move the image of the WHEEL
    tmrRepeat.Enabled = True 'Enables REPEAT function if key is held down
End Sub

'Release the RIGHT button
Private Sub imgRIGHT_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tmrRepeat.Enabled = False 'Disables the REPEAT function
    imgRIGHT.Move 80, 0 'Display the RIGHT button in the raised position
End Sub

'Press the LEFT button
Private Sub imgLEFT_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgLEFT.Move 2, 2 'Display the LEFT button in the pressed position
    Increment = -3 'Indicates the WHEEL is to be rotated by -3 pixels
    Call Wheel_Move 'Move the image of the WHEEL
    tmrRepeat.Enabled = True 'Enables REPEAT function if key is held down
End Sub

'Release the LEFT button
Private Sub imgLEFT_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tmrRepeat.Enabled = False 'Disables the REPEAT function
    imgLEFT.Move 0, 0 'Display the LEFT button in the raised position
End Sub

'If a BUTTON is held in the pressed postion the timer will repeat
'that BUTTON's function until the BUTTON is released.
Private Sub tmrRepeat_Timer()
    Call Wheel_Move
End Sub

'Detect if any of the ARROW keys are being pressed
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 37 ' LEFT arrow
            If m_Orientation = "H" Then
                Increment = -3 'Indicates the WHEEL is to be rotated by -3 pixels
                If m_ShowButtons Then imgLEFT.Move 2, 2 'Display the LEFT button in the pressed position
                Call Wheel_Move 'Move the image of the WHEEL
            End If
        Case 38 ' UP arrow
            If m_Orientation = "V" Then
                Increment = -3 'Indicates the WHEEL is to be rotated by -3 pixels
                If m_ShowButtons Then imgUP.Move 2, 2 'Display the UP button in the pressed position
                Call Wheel_Move 'Move the image of the WHEEL
            End If
        Case 39 ' RIGHT arrow
            If m_Orientation = "H" Then
                Increment = 3 'Indicates the WHEEL is to be rotated by +3 pixels
                If m_ShowButtons Then imgRIGHT.Move 82, 2 'Display the RIGHT button in the pressed position
                Call Wheel_Move 'Move the image of the WHEEL
            End If
        Case 40 ' DOWN arrow
            If m_Orientation = "V" Then
                Increment = 3 'Indicates the WHEEL is to be rotated by +3 pixels
                If m_ShowButtons Then imgDOWN.Move 2, 82 'Display the DOWN button in the pressed position
                Call Wheel_Move 'Move the image of the WHEEL
            End If
    End Select
End Sub

'Redisplay the BUTTON associated with the ARROW key being pressed
'in the raised position
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 37: If m_ShowButtons Then imgLEFT.Move 0, 0   'Display the LEFT  button in the raised position
        Case 38: If m_ShowButtons Then imgUP.Move 0, 0     'Display the UP    button in the raised position
        Case 39: If m_ShowButtons Then imgRIGHT.Move 80, 0 'Display the RIGHT button in the raised position
        Case 40: If m_ShowButtons Then imgDOWN.Move 0, 80  'Display the DOWN  button in the raised position
    End Select
End Sub

'GRAB the WHEEL
Private Sub picWHEEL_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    GRABBED = True 'Indicates that we have "GRABBED" the WHEEL
    LastX = x: LastY = y 'Holds the X-Y co-ordinates
End Sub

'Move the Mouse over the WHEEL
Private Sub picWHEEL_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If GRABBED Then 'Only if we have already "GRABBED" the WHEEL
        If m_Orientation = "V" Then 'If this is a VERTICAL control
            Increment = y - LastY 'Calculate the VERTICAL change in pixels
        Else
            Increment = x - LastX 'Calculate the HORIZONTAL change in pixels
        End If
        Call Wheel_Move 'Move the image of the WHEEL
        LastX = x: LastY = y 'Hold the new X-Y co-ordinates
    End If
End Sub

'Release the WHEEL
Private Sub picWHEEL_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    GRABBED = False
End Sub

'This routine calculates the WHEEL position, increments or decrements
'the VALUE of the control depending on the magnitude of the movement
'then calls the routine which actually draws the WHEEL.
Private Sub Wheel_Move()
    
    Dim intCntr As Integer 'Generic counter
    
    'No need to process if there's no change
    If Increment = 0 Then GoTo Wheel_Move_Exit
    'If MAX or MIN has been reached and "SpinOver" isn't
    'enabled then do no further processing
    If Not m_SpinOver Then
        Select Case Increment
            Case Is < 0
                If m_Value - 1 < m_Min Then GoTo Wheel_Move_Exit
            Case Is > 0
                If m_Value + 1 > m_Max Then GoTo Wheel_Move_Exit
        End Select
    End If
    
    'Loop once each time for the magnitude of the increment
    For intCntr = Sgn(Increment) To Increment Step Sgn(Increment)
        WheelPosition = WheelPosition + Sgn(Increment)
        'Upper and Lower Bounds of wheel position are 8 & 0
        If WheelPosition > 8 Then WheelPosition = 0
        If WheelPosition < 0 Then WheelPosition = 8
        Select Case WheelPosition
            'VALUE is only changed on every third movement
            Case 0, 3, 6
                m_Value = m_Value + Sgn(Increment) 'Increment the VALUE
                If Increment < 0 Then
                    If m_Value < m_Min Then
                        If m_SpinOver Then
                            m_Value = m_Max
                        Else
                            m_Value = m_Min
                            GoTo Wheel_Move_Exit
                        End If
                    End If
                    RaiseEvent Change
                    RaiseEvent SpinLess
                Else
                    If m_Value > m_Max Then
                        If m_SpinOver Then
                            m_Value = m_Min
                        Else
                            m_Value = m_Max
                            GoTo Wheel_Move_Exit
                        End If
                    End If
                    RaiseEvent Change
                    RaiseEvent SpinMore
                End If
        End Select
        DRAW_WHEEL
    Next intCntr

Wheel_Move_Exit:
    Exit Sub
End Sub

'This routine depicts the actual movement of the WHEEL
Private Sub DRAW_WHEEL()

    'StretchBlt is used to draw the image of the WHEEL from either
    'the Vertical or Horizontal originl image, only slices of one pixel
    'are taken at a time and are then stretched to 12 pixels.
    'The MASK colour is then "BitBlt"ed onto the image of the WHEEL
    'to give it some colour, this corresponds to the "ShadeWheel" setting.
    If m_Orientation = "V" Then
        StretchBlt picWHEEL.hdc, 0, 0, 12, 43, _
                   picV_SND.hdc, WheelPosition, 0, 1, 43, vbSrcCopy
        BitBlt picWHEEL.hdc, 0, 0, 12, 43, _
               picV_MSK.hdc, 0, 0, vbSrcAnd
    Else
        StretchBlt picWHEEL.hdc, 0, 0, 43, 12, _
                   picH_SND.hdc, 0, WheelPosition, 43, 1, vbSrcCopy
        BitBlt picWHEEL.hdc, 0, 0, 43, 12, _
               picH_MSK.hdc, 0, 0, vbSrcAnd
    End If
    'We now need to refresh the image as it's Autoredraw property
    'is set to True
    picWHEEL.Refresh

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    SHOW_ENABLED
End Property

Private Sub SHOW_ENABLED()

    If UserControl.Enabled Then
        shpWheelColour.DrawMode = 9
        shpWheelColour.FillColor = m_ShadeControl
        picWHEEL.Visible = True
    Else
        shpWheelColour.DrawMode = 15
        shpWheelColour.FillColor = &H808080
        picWHEEL.Visible = False
    End If

End Sub

Public Property Get Max() As Integer
Attribute Max.VB_Description = "Maximum Value that Wheel can attain."
Attribute Max.VB_ProcData.VB_Invoke_Property = "Data"
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Integer)
    Select Case New_Max
        Case Is < m_Value
            MsgBox "Max must not be less than Value", vbCritical, "SpinWheel"
        Case Is < m_Min
            MsgBox "Max must not be greater than Min", vbCritical, "SpinWheel"
        Case Else
            m_Max = New_Max
            PropertyChanged "Max"
    End Select
End Property

Public Property Get Min() As Integer
Attribute Min.VB_Description = "Minimum value that Wheel can attain."
Attribute Min.VB_ProcData.VB_Invoke_Property = "Data"
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Integer)
    Select Case New_Min
        Case Is > m_Value
            MsgBox "Min must not be greater than Value", vbCritical, "SpinWheel"
        Case Is > m_Max
            MsgBox "Min must not be greater than Max", vbCritical, "SpinWheel"
        Case Else
            m_Min = New_Min
            PropertyChanged "Min"
    End Select
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get Orientation() As String
Attribute Orientation.VB_Description = """V""ertical or ""H""orizontal"
Attribute Orientation.VB_ProcData.VB_Invoke_Property = "Colour"
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal New_Orientation As String)
    If Ambient.UserMode Then Err.Raise 393
    New_Orientation = UCase(New_Orientation)
    Select Case New_Orientation
        Case "H", "HORIZONTAL"
            m_Orientation = "H"
            PropertyChanged "Orientation"
            Call UserControl_Resize
        Case "V", "VERTICAL"
            m_Orientation = "V"
            PropertyChanged "Orientation"
            Call UserControl_Resize
        Case Else
            MsgBox "Value must be ""V"" or ""H""", vbCritical, "SpinWheel"
    End Select
End Property

Public Property Get ShadeControl() As OLE_COLOR
    ShadeControl = m_ShadeControl
End Property

Public Property Let ShadeControl(ByVal New_ShadeControl As OLE_COLOR)
    m_ShadeControl = New_ShadeControl
    PropertyChanged "ShadeControl"
    Call UserControl_Resize
End Property

Public Property Get ShadeWheel() As OLE_COLOR
    ShadeWheel = m_ShadeWheel
End Property

Public Property Let ShadeWheel(ByVal New_ShadeWheel As OLE_COLOR)
    m_ShadeWheel = New_ShadeWheel
    PropertyChanged "ShadeWheel"
    Call UserControl_Resize
End Property

Public Property Get ShowButtons() As Boolean
    ShowButtons = m_ShowButtons
End Property

Public Property Let ShowButtons(ByVal New_ShowButtons As Boolean)
    If Ambient.UserMode Then Err.Raise 393
    m_ShowButtons = New_ShowButtons
    PropertyChanged "ShowButtons"
    Call UserControl_Resize
End Property

Public Property Get SpinOver() As Boolean
    SpinOver = m_SpinOver
End Property

Public Property Let SpinOver(ByVal New_SpinOver As Boolean)
    m_SpinOver = New_SpinOver
    PropertyChanged "SpinOver"
End Property

Public Property Get Value() As Integer
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
    Select Case New_Value
        Case Is < m_Min
            MsgBox "Value must not be less than Min", vbCritical, "SpinWheel"
        Case Is > m_Max
            MsgBox "Value must not be greater than Max", vbCritical, "SpinWheel"
        Case Else
            Increment = 3 * (New_Value - m_Value)
            Call Wheel_Move
            m_Value = New_Value
            PropertyChanged "Value"
    End Select
End Property

'Initalize Variables when Control is activated
Private Sub UserControl_Initialize()
    WheelPosition = 0
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Max = m_def_Max
    m_Min = m_def_Min
    m_Orientation = m_def_Orientation
    m_ShadeControl = m_def_ShadeControl
    m_ShadeWheel = m_def_ShadeWheel
    m_ShowButtons = m_def_ShowButtons
    m_SpinOver = m_def_SpinOver
    m_Value = m_def_Value
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_SpinOver = PropBag.ReadProperty("SpinOver", m_def_SpinOver)
    m_ShowButtons = PropBag.ReadProperty("ShowButtons", m_def_ShowButtons)
    m_Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)
    m_ShadeWheel = PropBag.ReadProperty("ShadeWheel", m_def_ShadeWheel)
    m_ShadeControl = PropBag.ReadProperty("ShadeControl", m_def_ShadeControl)

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("SpinOver", m_SpinOver, m_def_SpinOver)
    Call PropBag.WriteProperty("ShowButtons", m_ShowButtons, m_def_ShowButtons)
    Call PropBag.WriteProperty("Orientation", m_Orientation, m_def_Orientation)
    Call PropBag.WriteProperty("ShadeWheel", m_ShadeWheel, m_def_ShadeWheel)
    Call PropBag.WriteProperty("ShadeControl", m_ShadeControl, m_def_ShadeControl)

End Sub

'The Dimensions of the Control are fixed depending on whether the
'property "ShowButtons" is set to true or false, and also on whether the
'property "Orientation" is set to "V"ertical or "H"orizontal.
'
'This Control uses Pixels as it's ScaleMode, however when refering to the
'dimensions of the Control itself it will always be measured in Twips,
'therefore I have represented these values in the format PIXELS * 15
'where 15 is the number of Twips per pixel.
'
Private Sub UserControl_Resize()
    'Set up the colours of the compenents with values read from
    'the property bag.
    shpWheelColour.FillColor = m_ShadeControl
    picV_MSK.BackColor = m_ShadeWheel
    picH_MSK.BackColor = m_ShadeWheel
    If m_Orientation = "V" Then
        'Hide Horizontal components and unhide Vertical components.
        imgHBack.Visible = False
        imgLEFT.Visible = False
        imgRIGHT.Visible = False
        imgVBack.Visible = True
        If UserControl.Width <> 28 * 15 Then UserControl.Width = 28 * 15
        If m_ShowButtons Then
            If UserControl.Height <> 101 * 15 Then UserControl.Height = 101 * 15
            imgVBack.Top = 21
            picWHEEL.Move 8, 29, 12, 43
            'Make the UP and DOWN buttons visible
            imgUP.Visible = True
            imgDOWN.Visible = True
        Else
            If UserControl.Height <> 59 * 15 Then UserControl.Height = 59 * 15
            imgVBack.Top = 0
            picWHEEL.Move 8, 8, 12, 43
            'Hide the UP and DOWN buttons
            imgUP.Visible = False
            imgDOWN.Visible = False
        End If
    Else
        'Hide Vertical components and unhide Horizontal components.
        imgHBack.Visible = True
        imgVBack.Visible = False
        imgUP.Visible = False
        imgDOWN.Visible = False
        If UserControl.Height <> 28 * 15 Then UserControl.Height = 28 * 15
        If m_ShowButtons Then
            If UserControl.Width <> 101 * 15 Then UserControl.Width = 101 * 15
            imgHBack.Left = 21
            picWHEEL.Move 29, 8, 43, 12
            'Make the LEFT and RIGHT buttons visible
            imgLEFT.Visible = True
            imgRIGHT.Visible = True
        Else
            If UserControl.Width <> 59 * 15 Then UserControl.Width = 59 * 15
            imgHBack.Left = 0
            picWHEEL.Move 8, 8, 43, 12
            'Hide the LEFT and RIGHT buttons
            imgLEFT.Visible = False
            imgRIGHT.Visible = False
        End If
    End If
    DRAW_WHEEL   'Draw the actual WHEEL
    SHOW_ENABLED 'Depict the control as Enabled or Disabled
End Sub


