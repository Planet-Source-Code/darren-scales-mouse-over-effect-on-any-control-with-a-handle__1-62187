VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   2580
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   900
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If IsOverhWnd(Command1.hwnd, x, y) Then
        Command1.Caption = "Over"
    Else
        Command1.Caption = "Command1"
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If IsOverhWnd(Picture1.hwnd, x, y) Then
        Picture1.BackColor = vbRed
    Else
        Picture1.BackColor = vbBlack
    End If
End Sub
