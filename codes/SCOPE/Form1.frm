VERSION 5.00
Object = "{BA920785-843E-4AA6-B237-72900B39C5BB}#3.0#0"; "KNOBSCONTROL.OCX"
Object = "{9ABA45C6-4143-11D4-879D-60CF5AC10000}#1.0#0"; "PRJKNOB.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7755
   ClientLeft      =   2520
   ClientTop       =   1260
   ClientWidth     =   12135
   ControlBox      =   0   'False
   DrawMode        =   7  'Invert
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7755
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton save 
      Caption         =   "save"
      Height          =   375
      Left            =   3000
      TabIndex        =   45
      Top             =   6840
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   8880
      TabIndex        =   41
      Text            =   "0"
      Top             =   6840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   8040
      TabIndex        =   40
      Text            =   "0"
      Top             =   6840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   120
   End
   Begin VB.CheckBox dual 
      Caption         =   "DUAL"
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6840
      Width           =   615
   End
   Begin VB.CheckBox add 
      Caption         =   "ADD"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   6840
      Width           =   615
   End
   Begin VB.CheckBox ch2 
      Caption         =   "CH2"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6840
      Width           =   615
   End
   Begin VB.CheckBox ch1 
      Caption         =   "CH1"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6840
      Width           =   615
   End
   Begin VB.CommandButton power_off 
      Caption         =   "off"
      Height          =   255
      Left            =   10680
      TabIndex        =   33
      Top             =   6960
      Width           =   735
   End
   Begin VB.TextBox show_time 
      Height          =   375
      Left            =   11160
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "0"
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox show_ch2 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "0"
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox show_ch1 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "0"
      Top             =   1080
      Width           =   615
   End
   Begin VB.CheckBox ch2_atten 
      Caption         =   "1/10"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6600
      Width           =   735
   End
   Begin VB.CheckBox slope 
      DownPicture     =   "Form1.frx":7207
      Height          =   615
      Left            =   10680
      Picture         =   "Form1.frx":7DA9
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5400
      Width           =   855
   End
   Begin VB.CheckBox ch2_filter 
      Caption         =   "DC"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6600
      Width           =   735
   End
   Begin VB.OptionButton trigger 
      Caption         =   "OneShot"
      Height          =   375
      Index           =   2
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4920
      Width           =   855
   End
   Begin VB.OptionButton trigger 
      Caption         =   "Auto"
      Height          =   375
      Index           =   1
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4560
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton trigger 
      Caption         =   "Normal"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4200
      Width           =   855
   End
   Begin VB.CheckBox ch1_atten 
      Caption         =   "1/10"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3000
      Width           =   735
   End
   Begin KnobOCXControl.Knob clock_selector 
      Height          =   1755
      Left            =   10200
      TabIndex        =   11
      Top             =   1440
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   3096
      MaximumValue    =   7
      BackColor       =   -2147483633
      MajorTicks      =   8
      MinorTicks      =   0
      ForceStdValue   =   -1  'True
      ForeColor       =   -2147483630
      Reg             =   "uy873kho2-34hkl5-56kj1"
      RememberValue   =   -1  'True
      UniqueID        =   8
   End
   Begin prjKnob.Knob intensity 
      Height          =   510
      Left            =   9000
      TabIndex        =   10
      Top             =   6720
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   900
      Max             =   255
   End
   Begin KnobOCXControl.Knob ch1_range 
      Height          =   1755
      Left            =   600
      TabIndex        =   9
      Top             =   1440
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   3096
      MaximumValue    =   7
      BackColor       =   12632256
      MajorTicks      =   8
      MinorTicks      =   0
      ForceStdValue   =   -1  'True
      ForeColor       =   -2147483630
      Reg             =   "uy873kho2-34hkl5-56kj1"
      RememberValue   =   -1  'True
      UniqueID        =   6
   End
   Begin prjKnob.Knob focus 
      Height          =   510
      Left            =   7800
      TabIndex        =   8
      Top             =   6720
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   900
      Value           =   5
   End
   Begin KnobOCXControl.Knob ch2_range 
      Height          =   1755
      Left            =   600
      TabIndex        =   7
      Top             =   4920
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   3096
      MaximumValue    =   7
      BackColor       =   12632256
      MajorTicks      =   8
      MinorTicks      =   0
      ForceStdValue   =   -1  'True
      ForeColor       =   -2147483630
      Reg             =   "uy873kho2-34hkl5-56kj1"
      RememberValue   =   -1  'True
      UniqueID        =   5
   End
   Begin VB.CheckBox ch1_filter 
      Caption         =   "DC"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox show_y 
      Height          =   375
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0"
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox show_x 
      Height          =   375
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0"
      Top             =   6480
      Width           =   495
   End
   Begin Project1.Wheel xpos 
      Height          =   420
      Left            =   5730
      TabIndex        =   2
      Top             =   5790
      Width           =   1515
      _ExtentX        =   1561
      _ExtentY        =   741
      Max             =   10
      Min             =   -10
      SpinOver        =   0   'False
      Orientation     =   "H"
   End
   Begin VB.PictureBox monitor 
      AutoRedraw      =   -1  'True
      DrawStyle       =   1  'Dash
      Height          =   4695
      Left            =   3480
      Picture         =   "Form1.frx":894B
      ScaleHeight     =   4635
      ScaleWidth      =   6075
      TabIndex        =   1
      Top             =   960
      Width           =   6140
   End
   Begin VB.CheckBox pause 
      Caption         =   "Pause"
      Height          =   255
      Left            =   9840
      Picture         =   "Form1.frx":160F8
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6960
      Width           =   735
   End
   Begin Project1.Wheel ypos 
      Height          =   1515
      Left            =   9705
      TabIndex        =   4
      Top             =   2550
      Width           =   420
      _ExtentX        =   1561
      _ExtentY        =   741
      Max             =   10
      Min             =   -10
      SpinOver        =   0   'False
      ShadeWheel      =   -2147483639
      ShadeControl    =   -2147483639
   End
   Begin VB.Label savelabel 
      Caption         =   "save picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   46
      Top             =   6600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label time 
      Caption         =   "time(ms)"
      Height          =   255
      Left            =   8280
      TabIndex        =   44
      Top             =   6600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label volt 
      Caption         =   "volt(v)"
      Height          =   255
      Left            =   9120
      TabIndex        =   43
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label measure 
      Caption         =   "Measure"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   42
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Power"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10200
      TabIndex        =   35
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   34
      Top             =   6555
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Trigger"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      TabIndex        =   32
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "time/div"
      Height          =   375
      Left            =   10440
      TabIndex        =   30
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "SWEEP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      TabIndex        =   29
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Y pos"
      Height          =   255
      Index           =   1
      Left            =   11040
      TabIndex        =   28
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "X pos"
      Height          =   375
      Index           =   0
      Left            =   6240
      TabIndex        =   27
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "intensity"
      Height          =   255
      Left            =   8400
      TabIndex        =   26
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Focus"
      Height          =   255
      Left            =   7320
      TabIndex        =   25
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "CH2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   24
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "volts/div"
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   23
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "volts/div"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   22
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "CH1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   21
      Top             =   600
      Width           =   615
   End
   Begin VB.Image Image6 
      Height          =   2880
      Left            =   10320
      Picture         =   "Form1.frx":23259
      Top             =   3480
      Width           =   1560
   End
   Begin VB.Image Image2 
      Height          =   885
      Left            =   2760
      Picture         =   "Form1.frx":25542
      Top             =   6480
      Width           =   3345
   End
   Begin VB.Image Image5 
      Height          =   2880
      Left            =   10320
      Picture         =   "Form1.frx":2701C
      Top             =   360
      Width           =   1560
   End
   Begin VB.Image Image4 
      Height          =   3405
      Left            =   480
      Picture         =   "Form1.frx":29305
      Top             =   3960
      Width           =   2130
   End
   Begin VB.Image Image3 
      Height          =   3405
      Left            =   480
      Picture         =   "Form1.frx":2AB79
      Top             =   360
      Width           =   2130
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   2760
      Picture         =   "Form1.frx":2C3ED
      Top             =   360
      Width           =   7500
   End
   Begin VB.Image Image7 
      Height          =   915
      Left            =   6960
      Picture         =   "Form1.frx":35559
      Top             =   6480
      Width           =   4785
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Dim command_b As Single
Dim command_c As Single
'''''''''''''''''''''''''''''''
Dim port_a As Single
Dim port_b As Single
Dim port_c As Single
Dim Control As Single
'''''''''''''''''''''''''''''''
Dim X, DownX, UpX As Single
Dim Y, DownY, UpY As Single
Dim frequency As Single
Dim down As Boolean

Private Sub add_Click()
dual.Value = False
End Sub

Private Sub ch1_atten_Click()
Dim atten_ch1 As Single
Dim atten_ch1_zoon As Single
Dim atten_ch1_mask As Single
''''''''''''''''''''''''''''''''''''
atten_ch1_zoon = 4
atten_ch1_mask = 251
''''''''''''''''''''''''''''''''''''
atten_ch1 = ch1_atten.Value * atten_ch1_zoon
command_c = command_c And atten_ch1_mask
command_c = command_c Or atten_ch1
Out port_c, command_c
End Sub
Private Sub ch1_Click()
dual.Value = False
End Sub
Private Sub ch1_filter_Click()
Dim filter_ch1 As Single
Dim filter_ch1_mask As Single
Dim filter_ch1_zoon As Single
'''''''''''''''''''''''''''
filter_ch1_zoon = 8
filter_ch1_mask = 495
'''''''''''''''''''''''''''
filter_ch1 = ch1_filter.Value * filter_ch1_zoon
command_c = command_c And filter_ch1_mask
command_c = command_c Or filter_ch1
Out port_c, command_c
End Sub
Private Sub ch1_range_Change()
Dim range_ch1 As Single
Dim atten1 As Single
Dim atten1_zoon As Single
Dim atten1_mask As Single
Dim atten_c As Single
Dim atten_c_mask As Single
''''''''''''''''''''''''''''
atten1_zoon = 128
atten1_mask = 127
atten_c_mask = 254
''''''''''''''''''''''''''''
range_ch1 = Round(ch1_range.Value)
atten1 = (range_ch1 And 1) * atten1_zoon
command_b = command_b And atten1_mask
command_b = command_b Or atten1
atten_c = range_ch1 And 6
command_c = (command_c And atten_c_mask) / 2
command_c = command_c Or atten_c
Out port_b, command_b
Out port_c, command_c
show_ch1.Text = range_ch1
End Sub
Private Sub ch2_Click()
dual.Value = False
End Sub
Private Sub monitor_MouseDown(Button As Integer, Shift As Integer, X1 As Single, Y1 As Single)
    
    DownX = X1
    DownY = Y1
   
End Sub

Private Sub monitor_MouseUp(Button As Integer, Shift As Integer, X2 As Single, Y2 As Single)
Dim mtime, mvolt As Single
 If (pause.Value = vbChecked) = True Then
    UpX = X2
    UpY = Y2

mtime = Round(Abs(UpX - DownX) / 7)
mvolt = Abs(UpY - DownY)
mtime = (mtime * 1000) / frequency
mtime = Round(mtime * 100) / 100
Text1.Text = mtime
End If
End Sub

Private Sub pause_click()
Text1.Visible = Not (Text1.Visible)
Text2.Visible = Not (Text2.Visible)
measure.Visible = Not (measure.Visible)
volt.Visible = Not (volt.Visible)
time.Visible = Not (time.Visible)
Label4.Visible = Not (Label4.Visible)
Label5.Visible = Not (Label5.Visible)
intensity.Visible = Not (intensity.Visible)
focus.Visible = Not (focus.Visible)
save.Visible = Not (save.Visible)
savelabel.Visible = Not (savelabel.Visible)
If (pause.Value = vbChecked) = True Then
monitor.MousePointer = 2
pause.Caption = "resume"
Else
monitor.MousePointer = 0
pause.Caption = "pause"
End If
Timer1.Enabled = Not (Timer1.Enabled)

End Sub

Private Sub dual_Click()
ch1.Value = False
ch2.Value = False
add.Value = False
End Sub


Private Sub intensity_Scroll()
monitor.ForeColor = 65536 * intensity.Value + 65280
End Sub
Private Sub focus_Scroll()
monitor.DrawWidth = 1 + focus.Value / 10
End Sub

Sub Form_Load()

monitor.ForeColor = vbGreen
monitor.DrawWidth = 1.5
port_a = 768
port_b = 769
port_c = 770
Control = 771
Out Control, 152
command_b = (Round(clock_selector.Value) * 4) Or ((Round(ch1_range.Value) And 1) * 128) Or (1)
command_c = (Round(ch1_range.Value) And 6)
Out port_b, command_b
Out port_c, command_c
End Sub

Private Sub power_off_Click()
Out port_b, 0
Out port_c, 0
End
End Sub


Private Sub slope_Click()
Dim slope_pol As Single
Dim slope_zoon As Single
Dim slope_mask As Single
'''''''''''''''''''''''''''''''''''
slope_zoon = 64
slope_mask = 191
'''''''''''''''''''''''''''''''''''
slope_pol = slope.Value * slope_zoon
command_b = command_b And slope_mask
command_b = command_b Or slope_pol
Out port_b, command_b
End Sub
Private Sub clock_selector_Change()
Dim clock_select As Single
Dim clock_zoon As Single
Dim clock_mask As Single
''''''''''''''''''''''''''''''''''
clock_zoon = 4
clock_mask = 227
''''''''''''''''''''''''''''''''''
Select Case clock_selector.Value
 Case 0
  frequency = 20000000
Case 1
  frequency = 10000000
Case 2
  frequency = 2000000
Case 3
  frequency = 1000000
Case 4
  frequency = 200000
Case 5
  frequency = 100000
Case 6
  frequency = 20000
Case 7
  frequency = 2000
  
End Select
clock_select = Round(clock_selector.Value) * clock_zoon
command_b = command_b And clock_mask
command_b = command_b Or clock_select
Out port_b, command_b
show_time.Text = clock_select
End Sub
Private Sub reset(reset As Single)
Dim reset_zoon As Single
Dim reset_mask As Single
''''''''''''''''''''''''''''
reset_zoon = 2
reset_mask = 253
''''''''''''''''''''''''''''
reset = reset * reset_zoon
command_b = command_b And reset_mask
command_b = command_b Or reset
End Sub
Private Sub start(a As Single)
Dim start_mask As Single
''''''''''''''''''''''''''''
start_mask = 254
''''''''''''''''''''''''''''
 command_b = command_b And start_mask
 command_b = command_b Or a
 Out port_b, command_b
End Sub
Private Sub inca()
Dim inca_zoon As Single
Dim inca_mask As Single
''''''''''''''''''''''''''''
inca_zoon = 32
inca_mask = 223
''''''''''''''''''''''''''''
command_b = command_b Or inca_zoon
Out port_b, command_b
command_b = command_b And inca_mask
Out port_b, command_b
End Sub

Private Sub wait_for_end_of_sampling()
Dim eos As Single
eos = 0
While (eos = 0)
eos = Inp(port_c)
eos = eos And 16
Wend
End Sub
Private Sub plot()
Dim sw As Single
Dim i As Single
Dim data1 As Single
Dim data2, data3 As Single
''''''''''''''''''''''
i = 1
X = xpos.Value * 50
monitor.CurrentX = X
sw = 0
data2 = Inp(port_a)
data3 = data2
While (i <= 1024)
data1 = Inp(port_a)
i = i + 1
If (Abs(data2 - data1) - Abs(data3 - data2)) < 1 Then
Y = 4250 - (17 * (data1 + ypos.Value))
End If
X = X + 7
monitor.Line -(X, Y)
 inca
 data3 = data2
 data2 = data1
 Wend
End Sub
Private Sub plot2()
Dim i As Single
Dim data As Single
''''''''''''''''''''''
i = 1
X = xpos.Value * 50
monitor.CurrentX = X
While (i <= 1024)
data1 = Inp(port_a)
i = i + 1
Y = 17 * (data1 + ypos.Value)
X = X + 7
monitor.Line -(X, Y)
 inca
Wend
End Sub
Private Sub Timer1_Timer()
  
  monitor.Cls
  monitor.CurrentY = Y
 If trigger(1).Value = True Then 'auto trigger mode
   reset (1)
   start (0)
   wait_for_end_of_sampling
   plot
   reset (0)
   start (1)
 End If
 If trigger(0).Value = True Then 'normal trigger mode
  wait_for_end_of_sampling
  plot2
 End If
 If trigger(2).Value = True Then 'one shot mode
 End If
End Sub

Private Sub xpos_Change()
show_x.Text = xpos.Value
End Sub

Private Sub ypos_Change()
show_y.Text = ypos.Value
End Sub

