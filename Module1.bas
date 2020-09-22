Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Public Function IsOverhWnd(hwnd As Long, X As Single, Y As Single) As Boolean
Dim Rec As RECT

    'get control position within the desktop
    If GetWindowRect(hwnd, Rec) = 0 Then Exit Function
    
    'x & y are currently in twips, so convert them to pixels
    X = X / Screen.TwipsPerPixelX
    Y = Y / Screen.TwipsPerPixelY
    
    'check if cursor is over the control
    If (X < 0) Or (Y < 0) Or (X > Rec.Right - Rec.Left) Or (Y > Rec.Bottom - Rec.Top) Then
        ReleaseCapture 'stop capturing the mouse
        IsOverhWnd = False
       Else
        SetCapture hwnd 'capture the mouse leaving the control
        IsOverhWnd = True
    End If
    
End Function
