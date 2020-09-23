Attribute VB_Name = "modTimer"
Option Explicit

Public Declare Function SetTimer Lib "user32" _
       (ByVal hwnd As Long, ByVal nIDEvent As Long, _
       ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" _
       (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
    
Public Const lElapse As Long = 1000 / 40 ' Holds API-Timer-Interval in milliseconds

Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, _
          ByVal idEvent As Long, ByVal dwTimer As Long)
          
    Static n As Integer
    
    ' draw the fire
    frmFire.ModifyPixels
    
    ' update CPULoad
    n = n + 1
    If n = 10 Then
        n = 0
        frmFire.DrawPGBar
    End If
    
End Sub



