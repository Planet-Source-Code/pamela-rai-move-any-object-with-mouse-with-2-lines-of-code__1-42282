VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   2700
      TabIndex        =   1
      Top             =   60
      Width           =   1740
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   330
      Left            =   555
      TabIndex        =   0
      Top             =   510
      Width           =   1485
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1410
      Left            =   1125
      Top             =   1395
      Width           =   2070
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'if you comment out movex or movey it will only move left/right or up/down
movey Button, Form1, Command1, False
movex Button, Form1, Command1
End Sub


Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
movey Button, Form1, Image1, False
movex Button, Form1, Image1
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
movey Button, Form1, List1, False
movex Button, Form1, List1
End Sub
