Attribute VB_Name = "modCCPbar"
' This module is used to place a ProgressBar in a StatusBar,
' and to modify Fore- and BackgroundColor of ProgressBar
' coded by phoenix 9/5/2004

Option Explicit

' API declarations
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Const WM_USER As Long = &H400
Private Const SB_GETRECT As Long = (WM_USER + 10)

' Set colors of progressbar
Public Sub SetProgressbarColor(ByVal hwnd As Long, _
  ByVal nForeColor As Long, _
  ByVal nBackColor As Long)

  ' set new foreground-color
  SendMessage hwnd, &H409, 0&, nForeColor

  ' set new background-color
  SendMessage hwnd, &H2001, 0&, nBackColor
End Sub

' Set only foreground-color of progressbar
Public Sub SetProgressbarForeColor(ByVal hwnd As Long, _
  ByVal nForeColor As Long)

  ' set new foreground-color
  SendMessage hwnd, &H409, 0&, nForeColor

End Sub


' Set ProgBar to StatusBar
Public Sub SetProgressBarToStatusBar( _
  ByVal hWnd_PBar As Long, _
  ByVal hWnd_SBar As Long, _
  ByVal nPanel As Long)

  Dim R As Rect
  
  ' Get size of panel
  SendMessageAny hWnd_SBar, SB_GETRECT, nPanel - 1, R

  ' lets give our progressbar a new home...
  SetParent hWnd_PBar, hWnd_SBar

  ' ... and set the correct position
  MoveWindow hWnd_PBar, R.Left, R.Top, R.Right - R.Left, _
    R.Bottom - R.Top, True
End Sub

