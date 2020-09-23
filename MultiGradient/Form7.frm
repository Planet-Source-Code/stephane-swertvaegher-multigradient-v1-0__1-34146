VERSION 5.00
Begin VB.Form Form7 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   Caption         =   "Form7"
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6900
   LinkTopic       =   "Form7"
   ScaleHeight     =   321
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4680
      Left            =   45
      Picture         =   "Form7.frx":0000
      ScaleHeight     =   312
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   450
      TabIndex        =   0
      Top             =   45
      Width           =   6750
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Picture1.Move 5, 5, Form7.ScaleWidth - 11, Form7.ScaleHeight - 11
T3D Form7, Picture1, 5, T3dRaiseInset
End Sub

Private Sub Picture1_Click()
Form7.Hide
End Sub
