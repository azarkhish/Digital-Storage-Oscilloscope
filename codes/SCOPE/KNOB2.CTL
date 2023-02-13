VERSION 5.00
Begin VB.UserControl Knob 
   ClientHeight    =   2010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2100
   ScaleHeight     =   2010
   ScaleWidth      =   2100
   Begin VB.PictureBox Knob1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "Knob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Value As Integer, oldX As Single, oldY As Single, K1 As New CLSKNOB
Dim P_Size As Integer, P_Mode As Integer, P_Step As Integer

Event Changed()
Sub SetBackColor(color As Long)
Knob1.BackColor = color
K1.drawKnob Knob1, Value, P_Mode, P_Step, P_Size
End Sub
Sub SetMode(NewMode As Integer)
P_Mode = NewMode
End Sub
Sub SetStep(NewStep As Integer)
P_Step = NewStep
End Sub
Sub SetVal(NewVal As Integer)
K1.SetValue P_Mode, NewVal
K1.drawKnob Knob1, Value, P_Mode, P_Step, P_Size
RaiseEvent Changed
End Sub
Function KnobValue() As Integer
KnobValue = K1.GetValue(P_Mode)
End Function

Private Sub Knob1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then

If Y < Knob1.ScaleHeight / 2 Then
If X > oldX Then Value = Value + 1
If X < oldX Then Value = Value - 1
Else
If X > oldX Then Value = Value - 1
If X < oldX Then Value = Value + 1
End If

If X < Knob1.ScaleWidth / 2 Then
If Y > oldY Then Value = Value - 1
If Y < oldY Then Value = Value + 1
Else
If Y > oldY Then Value = Value + 1
If Y < oldY Then Value = Value - 1
End If

K1.drawKnob Knob1, Value, P_Mode, P_Step, P_Size

RaiseEvent Changed

If Value > 100 Then Value = 50
If Value < -100 Then Value = -50

oldX = X
oldY = Y
End If
End Sub


Private Sub UserControl_Initialize()
P_Size = 0
P_Mode = 1
P_Step = 1
K1.BuildTable
K1.drawKnob Knob1, Value, P_Mode, P_Step, P_Size
End Sub

Private Sub UserControl_Resize()
Knob1.Height = UserControl.Height
Knob1.Width = UserControl.Width

Select Case Knob1.Height
Case Is <= 495: P_Size = 0
Case Is <= 615: P_Size = 1
Case Is <= 735: P_Size = 2
Case Is <= 855: P_Size = 3
Case Else: P_Size = 4
End Select

K1.drawKnob Knob1, Value, P_Mode, P_Step, P_Size
End Sub
