VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} chartClicker 
   Caption         =   "Chart Clicker"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2610
   OleObjectBlob   =   "chartClicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "chartClicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Public Declare Function SetTimer Lib "user32" ( _
'ByVal HWnd As Long, ByVal nIDEvent As Long, _
'ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long

'Public Declare Function KillTimer Lib "user32" ( _
'ByVal HWnd As Long, ByVal nIDEvent As Long) As Long

'Public TimerID As Long, TimerSeconds As Single, tim As Boolean

'Public Counter As Long

Private bStopped As Boolean


Sub Timer()
    If txtTimeLeft.Caption > 0 Then
        Application.Wait (Now + #12:00:01 AM#)
        txtTimeLeft.Caption = txtTimeLeft.Caption - 1
        Timer
    Else
        txtTimeLeft.Caption = txtTimeBetween.Value
        ExecuteGet
        Timer
    End If
        
End Sub



Sub ExecuteGet()
    If ActiveCell.Column > 1 Then
        Cells(ActiveCell.Row, 1).Select
   End If
   
   'trigger getbla using value in cell
   
    ActiveCell.Offset(1, 0).Select

End Sub



Private Sub btnAutomatic_Click()
    txtTimeLeft.Caption = txtTimeBetween.Value
    Timer
End Sub



Private Sub btnManual_Click()
    ExecuteGet
End Sub





