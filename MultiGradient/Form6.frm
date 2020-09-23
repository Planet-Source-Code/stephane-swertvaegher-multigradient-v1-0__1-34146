VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MultiGradient V1.0 "
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   398
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   512
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Done..."
      Height          =   375
      Left            =   6525
      TabIndex        =   3
      Top             =   5535
      Width           =   1050
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00D0D0D0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4920
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form6.frx":0000
      Top             =   495
      Width           =   7485
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   135
      Picture         =   "Form6.frx":0006
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   135
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Help me !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   7530
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Label7 = "Last action: search for help"
Form6.Hide
End Sub

Private Sub Form_Load()
ff = FreeFile
    On Error GoTo Load2
    Open App.Path & "\Help\Help.txt" For Input As #ff
    Text1.Text = Input(LOF(ff), 1)
    Close #ff
    Exit Sub
Load2:
    MsgBox "Something went wrong while loading help..." & vbCr & "Error " & Err.Number & vbCr & Err.Description, , MGTitle
End Sub

Private Sub Text1_GotFocus()
Picture1.SetFocus
End Sub
