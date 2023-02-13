VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLSKNOB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type PNT
X As Single
Y As Single
End Type

Dim TTS(359) As Double, TTC(359) As Double
Dim i As Integer, a As Single, lastVal As Integer, INN As Integer, OTN As Integer
Dim CenterPoint As PNT, P1 As PNT, P2 As PNT

Sub BuildTable()
For i = 0 To 359
TTS(i) = Sin(i * (3.14 / 180))
TTC(i) = Cos(i * (3.14 / 180))
Next
a = 315
End Sub

Sub drawKnob(PB As Object, Value As Integer, mode As Integer, step As Integer, size As Integer): On Error Resume Next

PB.AutoRedraw = True
PB.Cls
PB.DrawWidth = 2
PB.ScaleMode = 3

CenterPoint.X = PB.ScaleWidth / 2
CenterPoint.Y = PB.ScaleHeight / 2

Select Case Value
Case Is > lastVal
Select Case mode
Case 0: a = a - (27 * step)
Case 1: a = a - (2.7 * step)
Case 2: a = a - (0.27 * step)
End Select
Case Is < lastVal
Select Case mode
Case 0: a = a + (27 * step)
Case 1: a = a + (2.7 * step)
Case 2: a = a + (0.27 * step)
End Select
End Select

If a > 315 Then a = 315
If a < 45 Then a = 45

Select Case size
Case 0: INN = 6: OTN = 12
Case 1: INN = 8: OTN = 15
Case 2: INN = 9: OTN = 18
Case 3: INN = 12: OTN = 22
Case 4: INN = 15: OTN = 27
End Select

P1.X = CenterPoint.X + (INN * TTS(a))
P1.Y = CenterPoint.Y + (INN * TTC(a))
P2.X = CenterPoint.X + (OTN * TTS(a))
P2.Y = CenterPoint.Y + (OTN * TTC(a))

PB.Circle (CenterPoint.X, CenterPoint.Y), OTN
If size > 0 Then PB.Circle (CenterPoint.X + 0.5, CenterPoint.Y + 0.5), 0.8 * OTN Else PB.Circle (CenterPoint.X + 0.5, CenterPoint.Y + 0.5), 0.75 * OTN
PB.Line (P1.X, P1.Y)-(P2.X, P2.Y), &HFFFFFF

lastVal = Value
End Sub
Function GetValue(mode As Integer) As Integer
Select Case mode
Case 0: GetValue = Abs((a / 27) - 12)
Case 1: GetValue = Abs((a / 2.7) - 117)
Case 2: GetValue = Abs((a / 0.27) - 1167)
End Select
End Function

Sub SetValue(mode As Integer, val As Integer)
Select Case mode
Case 0: a = (-val + 12) * 27
Case 1: a = (-val + 117) * 2.7
Case 2: a = (-val + 1167) * 0.27
End Select
End Sub
