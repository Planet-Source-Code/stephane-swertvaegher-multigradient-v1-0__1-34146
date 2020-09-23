VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copy to clipboard"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6645
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   225
      Picture         =   "Form5.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   45
      Width           =   240
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   330
      Left            =   5760
      TabIndex        =   2
      Top             =   3150
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy data"
      Height          =   330
      Left            =   5265
      TabIndex        =   1
      Top             =   2700
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      Height          =   3120
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form5.frx":0442
      Top             =   405
      Width           =   4560
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Colors && Placement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   765
      TabIndex        =   4
      Top             =   45
      Width           =   2040
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Clipboard.Clear
Clipboard.SetText Text1.Text
MsgBox "The declarations & data are now stored on the clipboard" & vbCr & "Goto your VB-program, to the declarations section" & vbCr & "and hit Ctrl-C to paste it..." & vbCr & vbCr & "Note that the declarations are public, so they" & vbCr & "must be stored in a module...", , MGTitle

End Sub

Private Sub Command4_Click()
Form5.Hide
End Sub

Private Sub Form_Activate()
Picture1.SetFocus
Text1 = ""
Text1 = "'Declarations" & vbCrLf
Text1 = Text1 & "Public Mgrad&(9), MGPct!(9)" & vbCrLf
Text1 = Text1 & vbCrLf
Text1 = Text1 & "Public Sub MultiGradData" & vbCrLf
For xx = 0 To 9
Text1 = Text1 & "Mgrad(" & xx & ") = " & Str(Mgrad(xx)) & vbCrLf
Text1 = Text1 & "MgGPct(" & xx & ") = " & Str(MGPct(xx)) & vbCrLf
Next xx
Text1 = Text1 & "End Sub"
End Sub

Private Sub Text1_GotFocus()
Picture1.SetFocus
End Sub
