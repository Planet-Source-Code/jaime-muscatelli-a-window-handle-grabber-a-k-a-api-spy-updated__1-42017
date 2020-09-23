VERSION 5.00
Begin VB.Form FRMGRABBER 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window Grabber"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "FRMGRABBER.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picGrabber 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   240
      Picture         =   "FRMGRABBER.frx":12FA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   480
   End
   Begin VB.TextBox txtTExt 
      Height          =   1215
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox txtCLass 
      Height          =   1215
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label lblcaption 
      AutoSize        =   -1  'True
      Caption         =   "Caption:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   585
   End
   Begin VB.Label lblClass 
      AutoSize        =   -1  'True
      Caption         =   "Class Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   885
   End
End
Attribute VB_Name = "FRMGRABBER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal HWND As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal HWND As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal HWND As Long) As Long

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private MOUSE As POINTAPI



Private Sub picGrabber_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.MousePointer = 99
Me.MouseIcon = picGrabber.Picture
txtCLass.Text = vbNullString
txtTExt.Text = vbNullString
End Sub

Private Sub picGrabber_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Me.MousePointer = 1

GetCursorPos MOUSE

GETHWND MOUSE.x, MOUSE.y

End Sub

Private Sub GETHWND(x As Long, y As Long)
Dim lArray() As Long
Dim lFirstHwnd As Long
Dim lNumOfWin As Long
Dim I As Long

lFirstHwnd = WindowFromPoint(x, y)


While lFirstHwnd <> GetParent(lFirstHwnd)
    ReDim Preserve lArray(lNumOfWin)
    lArray(lNumOfWin) = lFirstHwnd
    lFirstHwnd = GetParent(lFirstHwnd)
    lNumOfWin = lNumOfWin + 1
Wend


For I = 0 To UBound(lArray)
    MAKECODE lArray(I)
Next I

End Sub
Private Sub MAKECODE(lhwnd As Long)
Dim sClassName As String * 50
Dim sHwndText As String * 256

GetClassName lhwnd, sClassName, 50
GetWindowText lhwnd, sHwndText, 256

If txtCLass.Text = vbNullString Then
txtCLass.Text = sClassName
txtTExt.Text = sHwndText
Exit Sub
End If
txtCLass.Text = txtCLass.Text & vbCrLf & sClassName
txtTExt.Text = txtTExt.Text & vbCrLf & sHwndText
End Sub

