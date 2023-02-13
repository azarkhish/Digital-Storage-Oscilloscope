VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMaxSpeedTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'
'The varaible to switch the timer on and
'     off... You have to switch it off some-ti
'     mes
Dim IsActive As Boolean
'
'This event will be generate
Public Event TikCall()


Public Property Get Enabled() As Boolean
    Enabled = IsActive
    Call Running
End Property


Public Property Let Enabled(ByVal SetEnabledTo As Boolean)
    IsActive = SetEnabledTo
    Call Running
End Property
'This function will generate your tikker
'


Private Sub Running()


    While IsActive
        RaiseEvent TikCall


        DoEvents
        Wend
    End Sub


Private Sub Class_Initialize()
    IsActive = False
    Call Running
End Sub


Private Sub Class_Terminate()
    IsActive = False
    Call Running
End Sub

