VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gradient rename"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   247
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   375
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2790
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   9
      Top             =   90
      Width           =   2715
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Rename"
      Height          =   330
      Left            =   4680
      TabIndex        =   6
      Top             =   2070
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done !"
      Height          =   330
      Left            =   4680
      TabIndex        =   5
      Top             =   2520
      Width           =   825
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00E0E0E0&
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
      Height          =   2820
      Left            =   90
      TabIndex        =   4
      Top             =   45
      Width           =   2625
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   2790
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1530
      Width           =   2715
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   90
      TabIndex        =   8
      Top             =   3330
      Width           =   5460
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   90
      TabIndex        =   7
      Top             =   2970
      Width           =   5460
   End
   Begin VB.Label Label3 
      Caption         =   "New gradient name:"
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
      Height          =   240
      Left            =   2835
      TabIndex        =   2
      Top             =   1260
      Width           =   2670
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
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
      Height          =   285
      Left            =   2835
      TabIndex        =   1
      Top             =   765
      Width           =   2670
   End
   Begin VB.Label label1 
      Caption         =   "Gradient name:"
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
      Height          =   240
      Left            =   2835
      TabIndex        =   0
      Top             =   495
      Width           =   2670
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldFile$, NewFile$

Private Sub Command1_Click()
Form1.Label7 = "Last action: Renaming gradients"
Form3.Hide
End Sub

Private Sub Command3_Click() 'rename
On Error GoTo RenameError
If Text1 = "" Then Exit Sub
If LCase(Right(Text1, 4)) <> ".mgr" Then
Text1 = Text1 & ".Mgr"
End If
NewFile = File1.Path & "\" & Text1
Label5 = NewFile
    For xx = 0 To File1.ListCount - 1
    If LCase(File1.List(xx)) = LCase(Text1) Then
    MsgBox "The gradient " & Text1 & " already exists !", , MGTitle
    Exit Sub
    End If
    Next xx
Name OldFile As NewFile
OldFile = NewFile
Label4 = OldFile
Label5 = NewFile
Label2 = Text1
File1.Refresh
Exit Sub
RenameError:
MsgBox Err.Number & vbCr & Err.Description, , MGTitle
End Sub

Private Sub File1_Click()
Label2 = File1.List(File1.ListIndex)
Text1 = Label2
Text1.SelStart = 0
Text1.SelLength = Len(Text1)
Text1.SetFocus
OldFile = File1.Path & "\" & Text1
Label4 = OldFile
NewFile = OldFile
Label5 = NewFile
ff = FreeFile
Open OldFile For Input As #1
For xx = 0 To 9
Input #ff, MGrad1(xx)
Input #ff, Mgpct3(xx)
MGPct1(xx) = Mgpct3(xx) / 10000
Next xx
Close #ff
Picture1.Cls
MultiGrad3 Picture1
End Sub

Private Sub Form_Activate()
On Error Resume Next
File1.Path = App.Path & "\Gradients"
File1.Refresh
Label2 = ""
Text1 = ""
Label4 = ""
File1.Selected(0) = True
End Sub

