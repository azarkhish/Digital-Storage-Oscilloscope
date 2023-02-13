VERSION 5.00
Begin VB.UserControl SwitchCheck 
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1530
   ScaleHeight     =   945
   ScaleWidth      =   1530
   Begin VB.Label Check1 
      BackStyle       =   0  'Transparent
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.Label LowerOffLbl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   0
      TabIndex        =   5
      Top             =   450
      Width           =   495
   End
   Begin VB.Label LowerOnLbl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please insert text here......."
      Height          =   855
      Left            =   600
      TabIndex        =   3
      Top             =   0
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Off"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   1
      Top             =   600
      Width           =   330
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   0
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "On"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "SwitchCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim IsChecked As Boolean
Dim IntChecked As Integer
Attribute IntChecked.VB_VarDescription = "Determine whether the switch is checked or not"
'Default Property Values:
Const m_def_ForeColor = &H80000012
Const m_def_FontTransparent = 0
Const m_def_WhatsThisHelpID = 0
Const m_def_Value = 0
Const m_def_UseMaskColor = 0
Const m_def_ToolTipText = ""
Const m_def_Style = 0
Const m_def_ScaleWidth = 0
Const m_def_ScaleTop = 0
Const m_def_ScaleMode = 0
Const m_def_ScaleLeft = 0
Const m_def_ScaleHeight = 0
Const m_def_RightToLeft = 0
Const m_def_PaletteMode = 0
Const m_def_OLEDropMode = 0
Const m_def_MousePointer = 0
Const m_def_MaskColor = 0
Const m_def_hWnd = 0
Const m_def_HitBehavior = 0
Const m_def_hDC = 0
Const m_def_HasDC = 0
'Const m_def_ForeColor = 0
'Const m_def_FontUnderline = 0
'Const m_def_FontTransparent = 0
'Const m_def_FontStrikethru = 0
'Const m_def_FontSize = 0
'Const m_def_FontName = ""
'Const m_def_FontItalic = 0
'Const m_def_FontBold = 0
Const m_def_FillStyle = 0
Const m_def_FillColor = 0
Const m_def_Enabled = 0
Const m_def_DrawWidth = 0
Const m_def_DrawStyle = 0
Const m_def_DrawMode = 0
Const m_def_Default = 0
Const m_def_CurrentY = 0
Const m_def_CurrentX = 0
Const m_def_ContainerHwnd = 0
Const m_def_ClipControls = 0
Const m_def_ClipBehavior = 0
Const m_def_CausesValidation = 0
'Const m_def_Caption = ""
Const m_def_Cancel = 0
Const m_def_BorderStyle = 0
'Const m_def_BackStyle = 0
Const m_def_BackColor = 0
Const m_def_AutoRedraw = 0
Const m_def_Appearance = 0
Const m_def_Checked = False
'Property Variables:
Dim m_ForeColor As OLE_COLOR
Dim m_FontTransparent As Boolean
Dim m_WhatsThisHelpID As Long
Dim m_Value As Boolean
Dim m_UseMaskColor As Boolean
Dim m_ToolTipText As String
Dim m_Style As Integer
Dim m_ScaleWidth As Single
Dim m_ScaleTop As Single
Dim m_ScaleMode As Integer
Dim m_ScaleLeft As Single
Dim m_ScaleHeight As Single
Dim m_RightToLeft As Boolean
Dim m_Picture As Picture
Dim m_PaletteMode As Integer
Dim m_Palette As Picture
Dim m_OLEDropMode As Integer
Dim m_MousePointer As Integer
Dim m_MouseIcon As Picture
Dim m_MaskPicture As Picture
Dim m_MaskColor As Long
Dim m_Image As Picture
Dim m_HyperLink As HyperLink
Dim m_hWnd As Long
Dim m_HitBehavior As Integer
Dim m_hDC As Long
Dim m_HasDC As Boolean
'Dim m_ForeColor As Long
'Dim m_FontUnderline As Boolean
'Dim m_FontTransparent As Boolean
'Dim m_FontStrikethru As Boolean
'Dim m_FontSize As Single
'Dim m_FontName As String
'Dim m_FontItalic As Boolean
'Dim m_FontBold As Boolean
'Dim m_Font As Font
Dim m_FillStyle As Integer
Dim m_FillColor As Long
Dim m_Enabled As Boolean
Dim m_DrawWidth As Integer
Dim m_DrawStyle As Integer
Dim m_DrawMode As Integer
Dim m_DownPicture As Picture
Dim m_DisabledPicture As Picture
Dim m_Default As Boolean
Dim m_DataMembers As DataMembers
Dim m_CurrentY As Single
Dim m_CurrentX As Single
Dim m_Controls As Object
Dim m_ContainerHwnd As Long
Dim m_ClipControls As Boolean
Dim m_ClipBehavior As Integer
Dim m_CausesValidation As Boolean
'Dim m_Caption As String
Dim m_Cancel As Boolean
Dim m_BorderStyle As Integer
'Dim m_BackStyle As Integer
Dim m_BackColor As Long
Dim m_AutoRedraw As Boolean
Dim m_Appearance As Integer
'Dim m_ActiveControl As Control
Dim m_Checked As Boolean
'Event Declarations:
Event LblClick() 'MappingInfo=Label3,Label3,-1,Click
Attribute LblClick.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event Click() 'MappingInfo=Check1,Check1,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=Check1,Check1,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
'Event WriteProperties(PropBag As PropertyBag)
Event Show()
Attribute Show.VB_Description = "Occurs when the control's Visible property changes to True."
Event Resize()
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
'Event ReadProperties(PropBag As PropertyBag)
Event Paint()
Attribute Paint.VB_Description = "Occurs when any part of a form or PictureBox control is moved, enlarged, or exposed."
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."
Event OLESetData(Data As DataObject, DataFormat As Integer)
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual."
Event OLECompleteDrag(Effect As Long)
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a manual or automatic drag/drop has been completed or canceled."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event InitProperties()
Attribute InitProperties.VB_Description = "Occurs the first time a user control or user document is created."
Event HitTest(X As Single, Y As Single, HitResult As Integer)
Attribute HitTest.VB_Description = "Occurs in a windowless user control in response to mouse activity."
Event Hide()
Attribute Hide.VB_Description = "Occurs when the control's Visible property changes to False."
Event GetDataMember(DataMember As String, Data As Object)
Attribute GetDataMember.VB_Description = "Occurs when a data consumer is asking this data source for one of it's data members."
'Event DblClick()
'Event Click()
'Event AsyncReadProgress(AsyncProp As AsyncProperty)
'Event AsyncReadComplete(AsyncProp As AsyncProperty)




Private Sub Check1_Click()
    RaiseEvent Click
If IsChecked = True Then IsChecked = False Else IsChecked = True
Call LoadCheck
End Sub

Public Sub LoadCheck()
    If IsChecked = True Then
        IntChecked = 1
        Label1.Visible = True
        Label2.Visible = False
        LowerOnLbl.Visible = True
        LowerOffLbl.Visible = False
    Else
        IntChecked = 0
        Label1.Visible = False
        Label2.Visible = True
        LowerOnLbl.Visible = False
        LowerOffLbl.Visible = True
    End If
End Sub

Private Sub UserControl_Initialize()
IsChecked = False
Checked = 0
Call LoadCheck
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get WhatsThisHelpID() As Long
Attribute WhatsThisHelpID.VB_Description = "Returns/sets an associated context number for an object."
    WhatsThisHelpID = m_WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal New_WhatsThisHelpID As Long)
    m_WhatsThisHelpID = New_WhatsThisHelpID
    PropertyChanged "WhatsThisHelpID"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Value() As Boolean
Attribute Value.VB_Description = "Returns/sets the value of an object."
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Boolean)
    m_Value = New_Value
    PropertyChanged "Value"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub ValidateControls()
Attribute ValidateControls.VB_Description = "Validate contents of the last control on the form before exiting the form"
     
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get UseMaskColor() As Boolean
Attribute UseMaskColor.VB_Description = "Returns or sets a value that determines whether the color assigned in the MaskColor property is used as a 'mask'. (That is, used to create transparent regions.)  Applies only if Style is set to 1."
    UseMaskColor = m_UseMaskColor
End Property

Public Property Let UseMaskColor(ByVal New_UseMaskColor As Boolean)
    m_UseMaskColor = New_UseMaskColor
    PropertyChanged "UseMaskColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = m_ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    m_ToolTipText = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12
Public Function TextWidth(ByVal Str As String) As Single
Attribute TextWidth.VB_Description = "Returns the width of a text string as it would be printed in the current font."

End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12
Public Function TextHeight(ByVal Str As String) As Single
Attribute TextHeight.VB_Description = "Returns the height of a text string as it would be printed in the current font."

End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Style() As Integer
Attribute Style.VB_Description = "Returns/sets the appearance of the control, whether standard (standard Windows style) or graphical (with a custom picture)."
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As Integer)
    m_Style = New_Style
    PropertyChanged "Style"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Size(ByVal Width As Single, ByVal Height As Single)
Attribute Size.VB_Description = "Changes the width and height of a User Control."
     
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12
Public Function ScaleY(ByVal Height As Single, Optional ByVal FromScale As Variant, Optional ByVal ToScale As Variant) As Single
Attribute ScaleY.VB_Description = "Converts the value for the height of a Form, PictureBox, or Printer from one unit of measure to another."

End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12
Public Function ScaleX(ByVal Width As Single, Optional ByVal FromScale As Variant, Optional ByVal ToScale As Variant) As Single
Attribute ScaleX.VB_Description = "Converts the value for the width of a Form, PictureBox, or Printer from one unit of measure to another."

End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
    ScaleWidth = m_ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
    m_ScaleWidth = New_ScaleWidth
    PropertyChanged "ScaleWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get ScaleTop() As Single
Attribute ScaleTop.VB_Description = "Returns/sets the vertical coordinates for the top edges of an object."
    ScaleTop = m_ScaleTop
End Property

Public Property Let ScaleTop(ByVal New_ScaleTop As Single)
    m_ScaleTop = New_ScaleTop
    PropertyChanged "ScaleTop"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get ScaleMode() As Integer
Attribute ScaleMode.VB_Description = "Returns/sets a value indicating measurement units for object coordinates when using graphics methods or positioning controls."
    ScaleMode = m_ScaleMode
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As Integer)
    m_ScaleMode = New_ScaleMode
    PropertyChanged "ScaleMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get ScaleLeft() As Single
Attribute ScaleLeft.VB_Description = "Returns/sets the horizontal coordinates for the left edges of an object."
    ScaleLeft = m_ScaleLeft
End Property

Public Property Let ScaleLeft(ByVal New_ScaleLeft As Single)
    m_ScaleLeft = New_ScaleLeft
    PropertyChanged "ScaleLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
    ScaleHeight = m_ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
    m_ScaleHeight = New_ScaleHeight
    PropertyChanged "ScaleHeight"
End Property

'The Underscore following "Scale" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Scale_(Optional X1 As Variant, Optional Y1 As Variant, Optional X2 As Variant, Optional Y2 As Variant)
    
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "Determines text display direction and control visual appearance on a bidirectional system."
    RightToLeft = m_RightToLeft
End Property

Public Property Let RightToLeft(ByVal New_RightToLeft As Boolean)
    m_RightToLeft = New_RightToLeft
    PropertyChanged "RightToLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     
End Sub

'The Underscore following "PSet" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub PSet_(X As Single, Y As Single, Color As Long)
    
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub PopupMenu(ByVal Menu As Object, Optional ByVal Flags As Variant, Optional ByVal X As Variant, Optional ByVal Y As Variant, Optional ByVal DefaultMenu As Variant)
Attribute PopupMenu.VB_Description = "Displays a pop-up menu on an MDIForm or Form object."
     
End Sub

'The Underscore following "Point" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8
Public Function Point(X As Single, Y As Single) As Long
Attribute Point.VB_Description = "Returns, as an integer of type Long, the RGB color of the specified point on a Form or PictureBox object."
    
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set m_Picture = New_Picture
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get PaletteMode() As Integer
Attribute PaletteMode.VB_Description = "Returns/sets a value that determines which palette to use for the controls on a object."
    PaletteMode = m_PaletteMode
End Property

Public Property Let PaletteMode(ByVal New_PaletteMode As Integer)
    m_PaletteMode = New_PaletteMode
    PropertyChanged "PaletteMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Palette() As Picture
Attribute Palette.VB_Description = "Returns/sets an image that contains the palette to use on an object when PaletteMode is set to Custom"
    Set Palette = m_Palette
End Property

Public Property Set Palette(ByVal New_Palette As Picture)
    Set m_Palette = New_Palette
    PropertyChanged "Palette"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub PaintPicture(ByVal Picture As Picture, ByVal X1 As Single, ByVal Y1 As Single, Optional ByVal Width1 As Variant, Optional ByVal Height1 As Variant, Optional ByVal X2 As Variant, Optional ByVal Y2 As Variant, Optional ByVal Width2 As Variant, Optional ByVal Height2 As Variant, Optional ByVal Opcode As Variant)
Attribute PaintPicture.VB_Description = "Draws the contents of a graphics file on a Form, PictureBox, or Printer object."
     
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get OLEDropMode() As Integer
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
    OLEDropMode = m_OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As Integer)
    m_OLEDropMode = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
     
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = m_MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    m_MousePointer = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = m_MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set m_MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get MaskPicture() As Picture
Attribute MaskPicture.VB_Description = "Returns/sets the picture that specifies the clickable/drawable area of the control when BackStyle is 0 (transparent)."
    Set MaskPicture = m_MaskPicture
End Property

Public Property Set MaskPicture(ByVal New_MaskPicture As Picture)
    Set m_MaskPicture = New_MaskPicture
    PropertyChanged "MaskPicture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get MaskColor() As Long
Attribute MaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the MaskPicture."
    MaskColor = m_MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As Long)
    m_MaskColor = New_MaskColor
    PropertyChanged "MaskColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Line(ByVal Flags As Integer, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal Color As Long)
Attribute Line.VB_Description = "Draws lines and rectangles on an object."
     
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Image() As Picture
Attribute Image.VB_Description = "Returns a handle, provided by Microsoft Windows, to a persistent bitmap."
    Set Image = m_Image
End Property

Public Property Set Image(ByVal New_Image As Picture)
    Set m_Image = New_Image
    PropertyChanged "Image"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=20,0,0,0
Public Property Get HyperLink() As HyperLink
Attribute HyperLink.VB_Description = "Returns a Hyperlink object used for browser style navigation."
    Set HyperLink = m_HyperLink
End Property

Public Property Set HyperLink(ByVal New_HyperLink As HyperLink)
    Set m_HyperLink = New_HyperLink
    PropertyChanged "HyperLink"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = m_hWnd
End Property

Public Property Let hWnd(ByVal New_hWnd As Long)
    m_hWnd = New_hWnd
    PropertyChanged "hWnd"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get HitBehavior() As Integer
Attribute HitBehavior.VB_Description = "Indicates which mode of automatic hit testing a windowless UserControl employs."
    HitBehavior = m_HitBehavior
End Property

Public Property Let HitBehavior(ByVal New_HitBehavior As Integer)
    m_HitBehavior = New_HitBehavior
    PropertyChanged "HitBehavior"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = m_hDC
End Property

Public Property Let hDC(ByVal New_hDC As Long)
    m_hDC = New_hDC
    PropertyChanged "hDC"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get HasDC() As Boolean
Attribute HasDC.VB_Description = "Determines whether a unique display context is allocated for the control."
    HasDC = m_HasDC
End Property

Public Property Let HasDC(ByVal New_HasDC As Boolean)
    m_HasDC = New_HasDC
    PropertyChanged "HasDC"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=8,0,0,0
'Public Property Get ForeColor() As Long
'    ForeColor = m_ForeColor
'End Property
'
'Public Property Let ForeColor(ByVal New_ForeColor As Long)
'    m_ForeColor = New_ForeColor
'    PropertyChanged "ForeColor"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=0,0,0,0
'Public Property Get FontUnderline() As Boolean
'    FontUnderline = m_FontUnderline
'End Property
'
'Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
'    m_FontUnderline = New_FontUnderline
'    PropertyChanged "FontUnderline"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=0,0,0,0
'Public Property Get FontTransparent() As Boolean
'    FontTransparent = m_FontTransparent
'End Property
'
'Public Property Let FontTransparent(ByVal New_FontTransparent As Boolean)
'    m_FontTransparent = New_FontTransparent
'    PropertyChanged "FontTransparent"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=0,0,0,0
'Public Property Get FontStrikethru() As Boolean
'    FontStrikethru = m_FontStrikethru
'End Property
'
'Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
'    m_FontStrikethru = New_FontStrikethru
'    PropertyChanged "FontStrikethru"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=12,0,0,0
'Public Property Get FontSize() As Single
'    FontSize = m_FontSize
'End Property
'
'Public Property Let FontSize(ByVal New_FontSize As Single)
'    m_FontSize = New_FontSize
'    PropertyChanged "FontSize"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=13,0,0,
'Public Property Get FontName() As String
'    FontName = m_FontName
'End Property
'
'Public Property Let FontName(ByVal New_FontName As String)
'    m_FontName = New_FontName
'    PropertyChanged "FontName"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=0,0,0,0
'Public Property Get FontItalic() As Boolean
'    FontItalic = m_FontItalic
'End Property
'
'Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
'    m_FontItalic = New_FontItalic
'    PropertyChanged "FontItalic"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=0,0,0,0
'Public Property Get FontBold() As Boolean
'    FontBold = m_FontBold
'End Property
'
'Public Property Let FontBold(ByVal New_FontBold As Boolean)
'    m_FontBold = New_FontBold
'    PropertyChanged "FontBold"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=6,0,0,0
'Public Property Get Font() As Font
'    Set Font = m_Font
'End Property
'
'Public Property Set Font(ByVal New_Font As Font)
'    Set m_Font = New_Font
'    PropertyChanged "Font"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get FillStyle() As Integer
Attribute FillStyle.VB_Description = "Returns/sets the fill style of a shape."
    FillStyle = m_FillStyle
End Property

Public Property Let FillStyle(ByVal New_FillStyle As Integer)
    m_FillStyle = New_FillStyle
    PropertyChanged "FillStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get FillColor() As Long
Attribute FillColor.VB_Description = "Returns/sets the color used to fill in shapes, circles, and boxes."
    FillColor = m_FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As Long)
    m_FillColor = New_FillColor
    PropertyChanged "FillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get DrawWidth() As Integer
Attribute DrawWidth.VB_Description = "Returns/sets the line width for output from graphics methods."
    DrawWidth = m_DrawWidth
End Property

Public Property Let DrawWidth(ByVal New_DrawWidth As Integer)
    m_DrawWidth = New_DrawWidth
    PropertyChanged "DrawWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get DrawStyle() As Integer
Attribute DrawStyle.VB_Description = "Determines the line style for output from graphics methods."
    DrawStyle = m_DrawStyle
End Property

Public Property Let DrawStyle(ByVal New_DrawStyle As Integer)
    m_DrawStyle = New_DrawStyle
    PropertyChanged "DrawStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get DrawMode() As Integer
Attribute DrawMode.VB_Description = "Sets the appearance of output from graphics methods or of a Shape or Line control."
    DrawMode = m_DrawMode
End Property

Public Property Let DrawMode(ByVal New_DrawMode As Integer)
    m_DrawMode = New_DrawMode
    PropertyChanged "DrawMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get DownPicture() As Picture
Attribute DownPicture.VB_Description = "Returns/sets a graphic to be displayed when the button is in the down position, if Style is set to 1."
    Set DownPicture = m_DownPicture
End Property

Public Property Set DownPicture(ByVal New_DownPicture As Picture)
    Set m_DownPicture = New_DownPicture
    PropertyChanged "DownPicture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get DisabledPicture() As Picture
Attribute DisabledPicture.VB_Description = "Returns/sets a graphic to be displayed when the button is disabled, if Style is set to 1."
    Set DisabledPicture = m_DisabledPicture
End Property

Public Property Set DisabledPicture(ByVal New_DisabledPicture As Picture)
    Set m_DisabledPicture = New_DisabledPicture
    PropertyChanged "DisabledPicture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Default() As Boolean
Attribute Default.VB_Description = "Determines which CommandButton control is the default command button on a form."
    Default = m_Default
End Property

Public Property Let Default(ByVal New_Default As Boolean)
    m_Default = New_Default
    PropertyChanged "Default"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=16,0,0,0
Public Property Get DataMembers() As DataMembers
Attribute DataMembers.VB_Description = "Returns a collection of data members to show at design time for this data source."
    Set DataMembers = m_DataMembers
End Property

Public Property Set DataMembers(ByVal New_DataMembers As DataMembers)
    Set m_DataMembers = New_DataMembers
    PropertyChanged "DataMembers"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub DataMemberChanged(ByVal DataMember As String)
Attribute DataMemberChanged.VB_Description = "Notify data consumers that a data member of this data source has changed."
     
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get CurrentY() As Single
Attribute CurrentY.VB_Description = "Returns/sets the vertical coordinates for next print or draw method."
    CurrentY = m_CurrentY
End Property

Public Property Let CurrentY(ByVal New_CurrentY As Single)
    m_CurrentY = New_CurrentY
    PropertyChanged "CurrentY"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get CurrentX() As Single
Attribute CurrentX.VB_Description = "Returns/sets the horizontal coordinates for next print or draw method."
    CurrentX = m_CurrentX
End Property

Public Property Let CurrentX(ByVal New_CurrentX As Single)
    m_CurrentX = New_CurrentX
    PropertyChanged "CurrentX"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=9,0,0,0
Public Property Get Controls() As Object
Attribute Controls.VB_Description = "A collection whose elements represent each control on a form, including elements of control arrays. "
    Set Controls = m_Controls
End Property

Public Property Set Controls(ByVal New_Controls As Object)
    Set m_Controls = New_Controls
    PropertyChanged "Controls"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ContainerHwnd() As Long
Attribute ContainerHwnd.VB_Description = "Returns a handle (from Microsoft Windows) to the window a UserControl is contained in."
    ContainerHwnd = m_ContainerHwnd
End Property

Public Property Let ContainerHwnd(ByVal New_ContainerHwnd As Long)
    m_ContainerHwnd = New_ContainerHwnd
    PropertyChanged "ContainerHwnd"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Cls()
Attribute Cls.VB_Description = "Clears graphics and text generated at run time from a Form, Image, or PictureBox."
     
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ClipControls() As Boolean
Attribute ClipControls.VB_Description = "Determines whether graphics methods in Paint events repaint an entire object or newly exposed areas."
    ClipControls = m_ClipControls
End Property

Public Property Let ClipControls(ByVal New_ClipControls As Boolean)
    m_ClipControls = New_ClipControls
    PropertyChanged "ClipControls"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get ClipBehavior() As Integer
Attribute ClipBehavior.VB_Description = "Indicates the manner in which a windowless UserControl's appearance is clipped."
    ClipBehavior = m_ClipBehavior
End Property

Public Property Let ClipBehavior(ByVal New_ClipBehavior As Integer)
    m_ClipBehavior = New_ClipBehavior
    PropertyChanged "ClipBehavior"
End Property

'The Underscore following "Circle" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Circle_(X As Single, Y As Single, Radius As Single, Color As Long, StartPos As Single, EndPos As Single, Aspect As Single)
    
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get CausesValidation() As Boolean
Attribute CausesValidation.VB_Description = "Returns/sets whether validation occurs on the control which lost focus."
    CausesValidation = m_CausesValidation
End Property

Public Property Let CausesValidation(ByVal New_CausesValidation As Boolean)
    m_CausesValidation = New_CausesValidation
    PropertyChanged "CausesValidation"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=13,0,0,
'Public Property Get Caption() As String
'    Caption = m_Caption
'End Property
'
'Public Property Let Caption(ByVal New_Caption As String)
'    m_Caption = New_Caption
'    PropertyChanged "Caption"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function CanPropertyChange(ByVal PropertyName As String) As Boolean
Attribute CanPropertyChange.VB_Description = "Asks the container if a property bound to a data source can be changed.  The CanPropertyChange method is most useful if the property specified in PropertyName is bound to a data source."

End Function
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=5
'Public Sub CancelAsyncRead(Optional ByVal Property As Variant)
'
'End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Cancel() As Boolean
Attribute Cancel.VB_Description = "Indicates whether a command button is the Cancel button on a form."
    Cancel = m_Cancel
End Property

Public Property Let Cancel(ByVal New_Cancel As Boolean)
    m_Cancel = New_Cancel
    PropertyChanged "Cancel"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=7,0,0,0
'Public Property Get BackStyle() As Integer
'    BackStyle = m_BackStyle
'End Property
'
'Public Property Let BackStyle(ByVal New_BackStyle As Integer)
'    m_BackStyle = New_BackStyle
'    PropertyChanged "BackStyle"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
    AutoRedraw = m_AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    m_AutoRedraw = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=5
'Public Sub AsyncRead(ByVal Target As String, ByVal AsyncType As Long, Optional ByVal PropertyName As Variant, Optional ByVal AsyncReadOptions As Variant)
'
'End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
    m_Appearance = New_Appearance
    PropertyChanged "Appearance"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=15,0,0,0
'Public Property Get ActiveControl() As Object
'    Set ActiveControl = m_ActiveControl
'End Property
'
'Public Property Set ActiveControl(ByVal New_ActiveControl As Control)
'    Set m_ActiveControl = New_ActiveControl
'    PropertyChanged "ActiveControl"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get Checked() As Boolean
    'Checked = m_Checked
    Checked = IsChecked
End Property

Public Property Let Checked(ByVal New_Checked As Boolean)
    m_Checked = New_Checked
    PropertyChanged "Checked"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_WhatsThisHelpID = m_def_WhatsThisHelpID
    m_Value = m_def_Value
    m_UseMaskColor = m_def_UseMaskColor
    m_ToolTipText = m_def_ToolTipText
    m_Style = m_def_Style
    m_ScaleWidth = m_def_ScaleWidth
    m_ScaleTop = m_def_ScaleTop
    m_ScaleMode = m_def_ScaleMode
    m_ScaleLeft = m_def_ScaleLeft
    m_ScaleHeight = m_def_ScaleHeight
    m_RightToLeft = m_def_RightToLeft
    Set m_Picture = LoadPicture("")
    m_PaletteMode = m_def_PaletteMode
    Set m_Palette = LoadPicture("")
    m_OLEDropMode = m_def_OLEDropMode
    m_MousePointer = m_def_MousePointer
    Set m_MouseIcon = LoadPicture("")
    Set m_MaskPicture = LoadPicture("")
    m_MaskColor = m_def_MaskColor
    Set m_Image = LoadPicture("")
    m_hWnd = m_def_hWnd
    m_HitBehavior = m_def_HitBehavior
    m_hDC = m_def_hDC
    m_HasDC = m_def_HasDC
'    m_ForeColor = m_def_ForeColor
'    m_FontUnderline = m_def_FontUnderline
'    m_FontTransparent = m_def_FontTransparent
'    m_FontStrikethru = m_def_FontStrikethru
'    m_FontSize = m_def_FontSize
'    m_FontName = m_def_FontName
'    m_FontItalic = m_def_FontItalic
'    m_FontBold = m_def_FontBold
'    Set m_Font = Ambient.Font
    m_FillStyle = m_def_FillStyle
    m_FillColor = m_def_FillColor
    m_Enabled = m_def_Enabled
    m_DrawWidth = m_def_DrawWidth
    m_DrawStyle = m_def_DrawStyle
    m_DrawMode = m_def_DrawMode
    Set m_DownPicture = LoadPicture("")
    Set m_DisabledPicture = LoadPicture("")
    m_Default = m_def_Default
    m_CurrentY = m_def_CurrentY
    m_CurrentX = m_def_CurrentX
    m_ContainerHwnd = m_def_ContainerHwnd
    m_ClipControls = m_def_ClipControls
    m_ClipBehavior = m_def_ClipBehavior
    m_CausesValidation = m_def_CausesValidation
'    m_Caption = m_def_Caption
    m_Cancel = m_def_Cancel
    m_BorderStyle = m_def_BorderStyle
'    m_BackStyle = m_def_BackStyle
    m_BackColor = m_def_BackColor
    m_AutoRedraw = m_def_AutoRedraw
    m_Appearance = m_def_Appearance
    m_Checked = m_def_Checked
    m_ForeColor = m_def_ForeColor
    m_FontTransparent = m_def_FontTransparent
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_WhatsThisHelpID = PropBag.ReadProperty("WhatsThisHelpID", m_def_WhatsThisHelpID)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_UseMaskColor = PropBag.ReadProperty("UseMaskColor", m_def_UseMaskColor)
    m_ToolTipText = PropBag.ReadProperty("ToolTipText", m_def_ToolTipText)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    m_ScaleWidth = PropBag.ReadProperty("ScaleWidth", m_def_ScaleWidth)
    m_ScaleTop = PropBag.ReadProperty("ScaleTop", m_def_ScaleTop)
    m_ScaleMode = PropBag.ReadProperty("ScaleMode", m_def_ScaleMode)
    m_ScaleLeft = PropBag.ReadProperty("ScaleLeft", m_def_ScaleLeft)
    m_ScaleHeight = PropBag.ReadProperty("ScaleHeight", m_def_ScaleHeight)
    m_RightToLeft = PropBag.ReadProperty("RightToLeft", m_def_RightToLeft)
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    m_PaletteMode = PropBag.ReadProperty("PaletteMode", m_def_PaletteMode)
    Set m_Palette = PropBag.ReadProperty("Palette", Nothing)
    m_OLEDropMode = PropBag.ReadProperty("OLEDropMode", m_def_OLEDropMode)
    m_MousePointer = PropBag.ReadProperty("MousePointer", m_def_MousePointer)
    Set m_MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    Set m_MaskPicture = PropBag.ReadProperty("MaskPicture", Nothing)
    m_MaskColor = PropBag.ReadProperty("MaskColor", m_def_MaskColor)
    Set m_Image = PropBag.ReadProperty("Image", Nothing)
    Set m_HyperLink = PropBag.ReadProperty("HyperLink", Nothing)
    m_hWnd = PropBag.ReadProperty("hWnd", m_def_hWnd)
    m_HitBehavior = PropBag.ReadProperty("HitBehavior", m_def_HitBehavior)
    m_hDC = PropBag.ReadProperty("hDC", m_def_hDC)
    m_HasDC = PropBag.ReadProperty("HasDC", m_def_HasDC)
'    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
'    m_FontUnderline = PropBag.ReadProperty("FontUnderline", m_def_FontUnderline)
'    m_FontTransparent = PropBag.ReadProperty("FontTransparent", m_def_FontTransparent)
'    m_FontStrikethru = PropBag.ReadProperty("FontStrikethru", m_def_FontStrikethru)
'    m_FontSize = PropBag.ReadProperty("FontSize", m_def_FontSize)
'    m_FontName = PropBag.ReadProperty("FontName", m_def_FontName)
'    m_FontItalic = PropBag.ReadProperty("FontItalic", m_def_FontItalic)
'    m_FontBold = PropBag.ReadProperty("FontBold", m_def_FontBold)
'    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_FillStyle = PropBag.ReadProperty("FillStyle", m_def_FillStyle)
    m_FillColor = PropBag.ReadProperty("FillColor", m_def_FillColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_DrawWidth = PropBag.ReadProperty("DrawWidth", m_def_DrawWidth)
    m_DrawStyle = PropBag.ReadProperty("DrawStyle", m_def_DrawStyle)
    m_DrawMode = PropBag.ReadProperty("DrawMode", m_def_DrawMode)
    Set m_DownPicture = PropBag.ReadProperty("DownPicture", Nothing)
    Set m_DisabledPicture = PropBag.ReadProperty("DisabledPicture", Nothing)
    m_Default = PropBag.ReadProperty("Default", m_def_Default)
    Set m_DataMembers = PropBag.ReadProperty("DataMembers", Nothing)
    m_CurrentY = PropBag.ReadProperty("CurrentY", m_def_CurrentY)
    m_CurrentX = PropBag.ReadProperty("CurrentX", m_def_CurrentX)
    Set m_Controls = PropBag.ReadProperty("Controls", Nothing)
    m_ContainerHwnd = PropBag.ReadProperty("ContainerHwnd", m_def_ContainerHwnd)
    m_ClipControls = PropBag.ReadProperty("ClipControls", m_def_ClipControls)
    m_ClipBehavior = PropBag.ReadProperty("ClipBehavior", m_def_ClipBehavior)
    m_CausesValidation = PropBag.ReadProperty("CausesValidation", m_def_CausesValidation)
'    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_Cancel = PropBag.ReadProperty("Cancel", m_def_Cancel)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
'    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_AutoRedraw = PropBag.ReadProperty("AutoRedraw", m_def_AutoRedraw)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
'    Set m_ActiveControl = PropBag.ReadProperty("ActiveControl", Nothing)
    m_Checked = PropBag.ReadProperty("Checked", m_def_Checked)
    Shape2.BackColor = PropBag.ReadProperty("OffColor", &HC000&)
    Shape1.BackColor = PropBag.ReadProperty("OnColor", &HC000&)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    Label3.FontName = PropBag.ReadProperty("FontName", "")
    Label3.FontSize = PropBag.ReadProperty("FontSize", 0)
    Label3.FontBold = PropBag.ReadProperty("FontBold", 0)
    Label3.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    Label3.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    Label3.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    m_FontTransparent = PropBag.ReadProperty("FontTransparent", m_def_FontTransparent)
    Set Label3.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Label3.Caption = PropBag.ReadProperty("Caption", "Please insert text here.......")
    Label3.BackColor = PropBag.ReadProperty("LblBackColor", &H8000000F)
    Label3.ForeColor = PropBag.ReadProperty("LblForeColor", &H80000012)
    Label3.AutoSize = PropBag.ReadProperty("AutoSize", False)
    Label3.WordWrap = PropBag.ReadProperty("WordWrap", False)
'    Label3.BackStyle = PropBag.ReadProperty("LblBackStyle", 1)
    Label3.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    Label3.BorderStyle = PropBag.ReadProperty("LblBorderStyle", 0)
End Sub

Private Sub UserControl_Resize()
Shape1.Height = UserControl.Height / 2
Shape2.Top = Shape1.Height
Shape2.Height = UserControl.Height / 2
Check1.Height = UserControl.Height
Check1.Width = 495
Label1.Top = Shape1.Height / 2 - 130
Label2.Top = Shape2.Top + Shape2.Height / 2 - 130
Label3.Height = UserControl.Height
Label3.Width = UserControl.Width - 600
Label3.Left = 600
LowerOnLbl.Top = Shape1.Top
LowerOffLbl.Top = Shape2.Top - 30
LowerOnLbl.Height = Shape1.Height
LowerOffLbl.Height = Shape2.Height + 15
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("WhatsThisHelpID", m_WhatsThisHelpID, m_def_WhatsThisHelpID)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("UseMaskColor", m_UseMaskColor, m_def_UseMaskColor)
    Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, m_def_ToolTipText)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("ScaleWidth", m_ScaleWidth, m_def_ScaleWidth)
    Call PropBag.WriteProperty("ScaleTop", m_ScaleTop, m_def_ScaleTop)
    Call PropBag.WriteProperty("ScaleMode", m_ScaleMode, m_def_ScaleMode)
    Call PropBag.WriteProperty("ScaleLeft", m_ScaleLeft, m_def_ScaleLeft)
    Call PropBag.WriteProperty("ScaleHeight", m_ScaleHeight, m_def_ScaleHeight)
    Call PropBag.WriteProperty("RightToLeft", m_RightToLeft, m_def_RightToLeft)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("PaletteMode", m_PaletteMode, m_def_PaletteMode)
    Call PropBag.WriteProperty("Palette", m_Palette, Nothing)
    Call PropBag.WriteProperty("OLEDropMode", m_OLEDropMode, m_def_OLEDropMode)
    Call PropBag.WriteProperty("MousePointer", m_MousePointer, m_def_MousePointer)
    Call PropBag.WriteProperty("MouseIcon", m_MouseIcon, Nothing)
    Call PropBag.WriteProperty("MaskPicture", m_MaskPicture, Nothing)
    Call PropBag.WriteProperty("MaskColor", m_MaskColor, m_def_MaskColor)
    Call PropBag.WriteProperty("Image", m_Image, Nothing)
    Call PropBag.WriteProperty("HyperLink", m_HyperLink, Nothing)
    Call PropBag.WriteProperty("hWnd", m_hWnd, m_def_hWnd)
    Call PropBag.WriteProperty("HitBehavior", m_HitBehavior, m_def_HitBehavior)
    Call PropBag.WriteProperty("hDC", m_hDC, m_def_hDC)
    Call PropBag.WriteProperty("HasDC", m_HasDC, m_def_HasDC)
'    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
'    Call PropBag.WriteProperty("FontUnderline", m_FontUnderline, m_def_FontUnderline)
'    Call PropBag.WriteProperty("FontTransparent", m_FontTransparent, m_def_FontTransparent)
'    Call PropBag.WriteProperty("FontStrikethru", m_FontStrikethru, m_def_FontStrikethru)
'    Call PropBag.WriteProperty("FontSize", m_FontSize, m_def_FontSize)
'    Call PropBag.WriteProperty("FontName", m_FontName, m_def_FontName)
'    Call PropBag.WriteProperty("FontItalic", m_FontItalic, m_def_FontItalic)
'    Call PropBag.WriteProperty("FontBold", m_FontBold, m_def_FontBold)
'    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("FillStyle", m_FillStyle, m_def_FillStyle)
    Call PropBag.WriteProperty("FillColor", m_FillColor, m_def_FillColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("DrawWidth", m_DrawWidth, m_def_DrawWidth)
    Call PropBag.WriteProperty("DrawStyle", m_DrawStyle, m_def_DrawStyle)
    Call PropBag.WriteProperty("DrawMode", m_DrawMode, m_def_DrawMode)
    Call PropBag.WriteProperty("DownPicture", m_DownPicture, Nothing)
    Call PropBag.WriteProperty("DisabledPicture", m_DisabledPicture, Nothing)
    Call PropBag.WriteProperty("Default", m_Default, m_def_Default)
    Call PropBag.WriteProperty("DataMembers", m_DataMembers, Nothing)
    Call PropBag.WriteProperty("CurrentY", m_CurrentY, m_def_CurrentY)
    Call PropBag.WriteProperty("CurrentX", m_CurrentX, m_def_CurrentX)
    Call PropBag.WriteProperty("Controls", m_Controls, Nothing)
    Call PropBag.WriteProperty("ContainerHwnd", m_ContainerHwnd, m_def_ContainerHwnd)
    Call PropBag.WriteProperty("ClipControls", m_ClipControls, m_def_ClipControls)
    Call PropBag.WriteProperty("ClipBehavior", m_ClipBehavior, m_def_ClipBehavior)
    Call PropBag.WriteProperty("CausesValidation", m_CausesValidation, m_def_CausesValidation)
'    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("Cancel", m_Cancel, m_def_Cancel)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
'    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("AutoRedraw", m_AutoRedraw, m_def_AutoRedraw)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
'    Call PropBag.WriteProperty("ActiveControl", m_ActiveControl, Nothing)
    Call PropBag.WriteProperty("Checked", m_Checked, m_def_Checked)
    Call PropBag.WriteProperty("OffColor", Shape2.BackColor, &HC000&)
    Call PropBag.WriteProperty("OnColor", Shape1.BackColor, &HC000&)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("FontName", Label3.FontName, "")
    Call PropBag.WriteProperty("FontSize", Label3.FontSize, 0)
    Call PropBag.WriteProperty("FontBold", Label3.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", Label3.FontItalic, 0)
    Call PropBag.WriteProperty("FontStrikethru", Label3.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontUnderline", Label3.FontUnderline, 0)
    Call PropBag.WriteProperty("FontTransparent", m_FontTransparent, m_def_FontTransparent)
    Call PropBag.WriteProperty("Font", Label3.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", Label3.Caption, "Please insert text here.......")
    Call PropBag.WriteProperty("LblBackColor", Label3.BackColor, &H8000000F)
    Call PropBag.WriteProperty("LblForeColor", Label3.ForeColor, &H80000012)
    Call PropBag.WriteProperty("AutoSize", Label3.AutoSize, False)
    Call PropBag.WriteProperty("WordWrap", Label3.WordWrap, False)
'    Call PropBag.WriteProperty("LblBackStyle", Label3.BackStyle, 1)
    Call PropBag.WriteProperty("BackStyle", Label3.BackStyle, 1)
    Call PropBag.WriteProperty("LblBorderStyle", Label3.BorderStyle, 0)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape2,Shape2,-1,BackColor
Public Property Get OffColor() As OLE_COLOR
Attribute OffColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    OffColor = Shape2.BackColor
End Property

Public Property Let OffColor(ByVal New_OffColor As OLE_COLOR)
    Shape2.BackColor() = New_OffColor
    PropertyChanged "OffColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape1,Shape1,-1,BackColor
Public Property Get OnColor() As OLE_COLOR
Attribute OnColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    OnColor = Shape1.BackColor
End Property

Public Property Let OnColor(ByVal New_OnColor As OLE_COLOR)
    Shape1.BackColor() = New_OnColor
    PropertyChanged "OnColor"
End Property

Private Sub Check1_DblClick()
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&h80000012&
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label3,Label3,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = Label3.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    Label3.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label3,Label3,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = Label3.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    Label3.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label3,Label3,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = Label3.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    Label3.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label3,Label3,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = Label3.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    Label3.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label3,Label3,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
    FontStrikethru = Label3.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    Label3.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label3,Label3,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
    FontUnderline = Label3.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    Label3.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get FontTransparent() As Boolean
Attribute FontTransparent.VB_Description = "Returns/sets a value that determines whether background text/graphics on a Form, Printer or PictureBox are displayed."
    FontTransparent = m_FontTransparent
End Property

Public Property Let FontTransparent(ByVal New_FontTransparent As Boolean)
    m_FontTransparent = New_FontTransparent
    PropertyChanged "FontTransparent"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label3,Label3,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Label3.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label3.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label3,Label3,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = Label3.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label3.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label3,Label3,-1,BackColor
Public Property Get LblBackColor() As OLE_COLOR
Attribute LblBackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    LblBackColor = Label3.BackColor
End Property

Public Property Let LblBackColor(ByVal New_LblBackColor As OLE_COLOR)
    Label3.BackColor() = New_LblBackColor
    PropertyChanged "LblBackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label3,Label3,-1,ForeColor
Public Property Get LblForeColor() As OLE_COLOR
Attribute LblForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    LblForeColor = Label3.ForeColor
End Property

Public Property Let LblForeColor(ByVal New_LblForeColor As OLE_COLOR)
    Label3.ForeColor() = New_LblForeColor
    PropertyChanged "LblForeColor"
End Property

Private Sub Label3_Click()
    RaiseEvent LblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label3,Label3,-1,AutoSize
Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Determines whether a control is automatically resized to display its entire contents."
    AutoSize = Label3.AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    Label3.AutoSize() = New_AutoSize
    PropertyChanged "AutoSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label3,Label3,-1,WordWrap
Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "Returns/sets a value that determines whether a control expands to fit the text in its Caption."
    WordWrap = Label3.WordWrap
End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
    Label3.WordWrap() = New_WordWrap
    PropertyChanged "WordWrap"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Label3,Label3,-1,BackStyle
'Public Property Get LblBackStyle() As Integer
'    LblBackStyle = Label3.BackStyle
'End Property
'
'Public Property Let LblBackStyle(ByVal New_LblBackStyle As Integer)
'    Label3.BackStyle() = New_LblBackStyle
'    PropertyChanged "LblBackStyle"
'End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label3,Label3,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = Label3.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    Label3.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label3,Label3,-1,BorderStyle
Public Property Get LblBorderStyle() As Integer
Attribute LblBorderStyle.VB_Description = "Returns/sets the border style for an object."
    LblBorderStyle = Label3.BorderStyle
End Property

Public Property Let LblBorderStyle(ByVal New_LblBorderStyle As Integer)
    Label3.BorderStyle() = New_LblBorderStyle
    PropertyChanged "LblBorderStyle"
End Property

