VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Moving Transparency"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1320
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1320
      ScaleWidth      =   4350
      TabIndex        =   0
      Top             =   0
      Width           =   4350
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Just Give It A Second... It Will Start."
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   3855
      End
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1320
      Left            =   -4400
      Picture         =   "Form1.frx":12C02
      ScaleHeight     =   1320
      ScaleWidth      =   4350
      TabIndex        =   1
      Top             =   0
      Width           =   4350
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" (ByVal hwnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Option Explicit
Private hRgn As Long
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_PATHMUSTEXIST = &H800
Private Const CC_FULLOPEN = &H2
Private Const CC_SOLIDCOLOR = &H80
Private Const CC_RGBINIT = &H1
Private Const CC_ANYCOLOR = &H100
Private Sub SetRegion()
If hRgn Then DeleteObject hRgn
hRgn = GetBitmapRegion(Picture1.Picture, vbRed)
SetWindowRgn Picture1.hwnd, hRgn, True
End Sub




Private Sub SetRegionTwo()
If hRgn Then DeleteObject hRgn
hRgn = GetBitmapRegion(Picture2.Picture, vbWhite)
SetWindowRgn Picture2.hwnd, hRgn, True
End Sub

Private Sub Form_Load()
SetRegion
SetRegionTwo
Show
Do
Picture2.Left = Picture2.Left + 10
Pause (0.001)
Loop Until Picture2.Left > Form1.Left + Form1.Width
MsgBox "That's It!"
End
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub


Private Sub Form_Terminate()
End
End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Private Sub Label1_Click()

End Sub


