VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.UserControl LED 
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   300
   ScaleHeight     =   20
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   20
   ToolboxBitmap   =   "LED.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1020
      Top             =   1500
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   300
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2100
      Top             =   1500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   18
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LED.ctx":0312
            Key             =   "RedCirLo"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LED.ctx":0814
            Key             =   "RedCirHi"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LED.ctx":0D16
            Key             =   "GreCirLo"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LED.ctx":1218
            Key             =   "GreCirHi"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LED.ctx":171A
            Key             =   "RedSqeLo"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LED.ctx":1C1C
            Key             =   "RedSqeHi"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LED.ctx":211E
            Key             =   "GreSqeLo"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LED.ctx":2620
            Key             =   "GreSqeHi"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LED.ctx":2B22
            Key             =   "YelCirLo"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LED.ctx":3024
            Key             =   "YelCirHi"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LED.ctx":3526
            Key             =   "YelSqeLo"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LED.ctx":3A28
            Key             =   "YelSqeHi"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LED.ctx":3F2A
            Key             =   "ButtonUp"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LED.ctx":55A8
            Key             =   "ButtonDown"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LED.ctx":6C26
            Key             =   "KeyOff"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LED.ctx":A2D8
            Key             =   "KeyOn"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LED.ctx":D98A
            Key             =   "KeyOffwo"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LED.ctx":F2F4
            Key             =   "KeyOnwo"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "LED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Default Property Values:
Const m_def_ButtonMode = 0
Const m_def_LEDColor = 0
Const m_def_Shape = 0
Const m_def_Value = False
'Property Variables:
Dim m_ButtonMode As enButtMode
Dim m_LEDColor As enLedColor
Dim m_Shape As LedShape
Dim m_Value As Boolean
'Event Declarations:
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Image1,Image1,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Image1,Image1,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Click() 'MappingInfo=Image1,Image1,-1,Click

Enum LedShape
    ledCircle = 0
    ledSquare = 1
    ledButton = 2
End Enum

Enum enLedColor
    ledRed = 0
    ledGreen = 1
    ledYellow = 2
End Enum

Enum enButtMode
    ledPushButton = 0
    ledSwitch = 1
End Enum


Public Property Get LEDColor() As enLedColor
Attribute LEDColor.VB_Description = "Color of LED"
Attribute LEDColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute LEDColor.VB_MemberFlags = "200"
    LEDColor = m_LEDColor
End Property

Public Property Let LEDColor(ByVal New_LEDColor As enLedColor)
    m_LEDColor = New_LEDColor
    PropertyChanged "LEDColor"
    Call Nastavi
End Property

Public Property Get Shape() As LedShape
Attribute Shape.VB_Description = "Shape of LED"
Attribute Shape.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Shape = m_Shape
End Property

Public Property Let Shape(ByVal New_Shape As LedShape)
    
    m_Shape = New_Shape
    PropertyChanged "Shape"
    Call Nastavi
End Property

Public Property Get Value() As Boolean
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Data"
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Boolean)
    m_Value = New_Value
    PropertyChanged "Value"
    Call Nastavi
End Property

Private Sub Timer1_Timer()
Static test As Boolean
    If m_Shape = 0 Then
        If m_LEDColor = 0 Then
            If test = False Then
                Image1.Picture = ImageList1.ListImages("RedCirLo").Picture
                test = True
            Else
                Image1.Picture = ImageList1.ListImages("RedCirHi").Picture
                test = False
            End If
        ElseIf m_LEDColor = 1 Then
            If test = False Then
                Image1.Picture = ImageList1.ListImages("GreCirLo").Picture
                test = True
            Else
                Image1.Picture = ImageList1.ListImages("GreCirHi").Picture
                test = False
            End If
        Else
            If test = False Then
                Image1.Picture = ImageList1.ListImages("YelCirLo").Picture
                test = True
            Else
                Image1.Picture = ImageList1.ListImages("YelCirHi").Picture
                test = False
            End If
        End If
    ElseIf m_Shape = 1 Then
        If m_LEDColor = 0 Then
            If test = False Then
                Image1.Picture = ImageList1.ListImages("RedSqeLo").Picture
                test = True
            Else
                Image1.Picture = ImageList1.ListImages("RedSqeHi").Picture
                test = False
            End If
        ElseIf m_LEDColor = 1 Then
            If test = False Then
                Image1.Picture = ImageList1.ListImages("GreSqeLo").Picture
                test = True
            Else
                Image1.Picture = ImageList1.ListImages("GreSqeHi").Picture
                test = False
            End If
        Else
            If test = False Then
                Image1.Picture = ImageList1.ListImages("YelSqeLo").Picture
                test = True
            Else
                Image1.Picture = ImageList1.ListImages("YelSqeHi").Picture
                test = False
            End If
        End If
    End If

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_LEDColor = m_def_LEDColor
    m_Shape = m_def_Shape
    m_Value = m_def_Value
    m_ButtonMode = m_def_ButtonMode
    Call Nastavi
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_LEDColor = PropBag.ReadProperty("LEDColor", m_def_LEDColor)
    m_Shape = PropBag.ReadProperty("Shape", m_def_Shape)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    Timer1.Enabled = PropBag.ReadProperty("Blink", False)
    Timer1.Interval = PropBag.ReadProperty("BlinkInterval", 1000)
    m_ButtonMode = PropBag.ReadProperty("ButtonMode", m_def_ButtonMode)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Resize()
Image1.Width = UserControl.ScaleWidth
Image1.Height = UserControl.ScaleHeight
End Sub

Private Sub UserControl_Show()
    Call Nastavi
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("LEDColor", m_LEDColor, m_def_LEDColor)
    Call PropBag.WriteProperty("Shape", m_Shape, m_def_Shape)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Blink", Timer1.Enabled, False)
    Call PropBag.WriteProperty("BlinkInterval", Timer1.Interval, 1000)
    Call PropBag.WriteProperty("ButtonMode", m_ButtonMode, m_def_ButtonMode)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub


Public Sub Nastavi()
    If m_Shape = 0 Then
        If m_LEDColor = 0 Then
            If m_Value = False Then
                Image1.Picture = ImageList1.ListImages("RedCirLo").Picture
            Else
                Image1.Picture = ImageList1.ListImages("RedCirHi").Picture
            End If
        ElseIf m_LEDColor = 1 Then
            If m_Value = False Then
                Image1.Picture = ImageList1.ListImages("GreCirLo").Picture
            Else
                Image1.Picture = ImageList1.ListImages("GreCirHi").Picture
            End If
        Else
            If m_Value = False Then
                Image1.Picture = ImageList1.ListImages("YelCirLo").Picture
            Else
                Image1.Picture = ImageList1.ListImages("YelCirHi").Picture
            End If
        End If
    ElseIf m_Shape = 1 Then
        If m_LEDColor = 0 Then
            If m_Value = False Then
                Image1.Picture = ImageList1.ListImages("RedSqeLo").Picture
            Else
                Image1.Picture = ImageList1.ListImages("RedSqeHi").Picture
            End If
        ElseIf m_LEDColor = 1 Then
            If m_Value = False Then
                Image1.Picture = ImageList1.ListImages("GreSqeLo").Picture
            Else
                Image1.Picture = ImageList1.ListImages("GreSqeHi").Picture
            End If
        Else
            If m_Value = False Then
                Image1.Picture = ImageList1.ListImages("YelSqeLo").Picture
            Else
                Image1.Picture = ImageList1.ListImages("YelSqeHi").Picture
            End If
        End If
    ElseIf m_Shape = 2 Then
        If m_Value = False Then
            Image1.Picture = ImageList1.ListImages("KeyOffwo").Picture
        Else
            Image1.Picture = ImageList1.ListImages("KeyOnwo").Picture
        End If
    End If
End Sub
            
            
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Timer1,Timer1,-1,Enabled
Public Property Get Blink() As Boolean
Attribute Blink.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Blink.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Blink = Timer1.Enabled
End Property

Public Property Let Blink(ByVal New_Blink As Boolean)
    Timer1.Enabled() = New_Blink
    PropertyChanged "Blink"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Timer1,Timer1,-1,Interval
Public Property Get BlinkInterval() As Long
Attribute BlinkInterval.VB_Description = "Returns/sets the number of milliseconds between calls to a Timer control's Timer event."
Attribute BlinkInterval.VB_ProcData.VB_Invoke_Property = ";Behavior"
    BlinkInterval = Timer1.Interval
End Property

Public Property Let BlinkInterval(ByVal New_BlinkInterval As Long)
    Timer1.Interval() = New_BlinkInterval
    PropertyChanged "BlinkInterval"
End Property
Private Sub Image1_Click()
If m_ButtonMode = ledSwitch Then
    If m_Shape = ledButton Then
        If m_Value = True Then
            Image1.Picture = ImageList1.ListImages("KeyOffwo").Picture
            m_Value = False
            PropertyChanged "Value"
         Else
            Image1.Picture = ImageList1.ListImages("KeyOnwo").Picture
            m_Value = True
            PropertyChanged "Value"
         End If
    End If
End If
RaiseEvent Click
End Sub

Public Property Get ButtonMode() As enButtMode
Attribute ButtonMode.VB_ProcData.VB_Invoke_Property = ";Misc"
    ButtonMode = m_ButtonMode
End Property

Public Property Let ButtonMode(ByVal New_ButtonMode As enButtMode)
    m_ButtonMode = New_ButtonMode
    PropertyChanged "ButtonMode"
End Property

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If m_ButtonMode = ledPushButton Then
    If m_Shape = ledButton Then
        Image1.Picture = ImageList1.ListImages("KeyOnwo").Picture
        m_Value = True
        PropertyChanged "Value"
    End If
End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If m_ButtonMode = ledPushButton Then
    If m_Shape = ledButton Then
        Image1.Picture = ImageList1.ListImages("KeyOffwo").Picture
        m_Value = False
        PropertyChanged "Value"
    End If
End If
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

