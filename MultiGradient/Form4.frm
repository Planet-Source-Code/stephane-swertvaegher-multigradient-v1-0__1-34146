VERSION 5.00
Begin VB.Form Form4 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "MultiGradient V1.0"
   ClientHeight    =   7965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7875
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   531
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7800
      Left            =   90
      ScaleHeight     =   520
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   520
      TabIndex        =   0
      Top             =   90
      Width           =   7800
      Begin VB.CommandButton Command7 
         Caption         =   "Diag. Rev."
         Height          =   330
         Left            =   1305
         TabIndex        =   7
         Top             =   855
         Width           =   1050
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Vert. Rev."
         Height          =   330
         Left            =   1305
         TabIndex        =   6
         Top             =   495
         Width           =   1050
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Hor. Rev."
         Height          =   330
         Left            =   1305
         TabIndex        =   5
         Top             =   135
         Width           =   1050
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Diagonal"
         Height          =   330
         Left            =   180
         TabIndex        =   4
         Top             =   855
         Width           =   1050
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         Height          =   330
         Left            =   3060
         TabIndex        =   3
         Top             =   135
         Width           =   690
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Vertical"
         Height          =   330
         Left            =   180
         TabIndex        =   2
         Top             =   495
         Width           =   1050
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Horizontal"
         Height          =   330
         Left            =   180
         TabIndex        =   1
         Top             =   135
         Width           =   1050
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() ' hor
MultiGrad Pic1, 0, False
End Sub

Private Sub Command5_Click() 'hor reversed
MultiGrad Pic1, 0, True
End Sub

Private Sub Command2_Click() 'vert
MultiGrad Pic1, 1, False
End Sub

Private Sub Command6_Click() 'vert rev
MultiGrad Pic1, 1, True
End Sub

Private Sub Command4_Click()
MultiGrad Pic1, 2, False
End Sub

Private Sub Command7_Click()
MultiGrad Pic1, 2, True
End Sub

Private Sub Command3_Click()
Form1.Label7 = "Last action: view gradient"
Form4.Hide
End Sub

Private Sub Form_Load()
Form4.Width = 8100
Form4.Height = 8100
Pic1.Move 5, 5, Form4.ScaleWidth - 11, Form4.ScaleHeight - 11
T3D Form4, Pic1, 4, T3dRaiseInset, T3dF1
End Sub

