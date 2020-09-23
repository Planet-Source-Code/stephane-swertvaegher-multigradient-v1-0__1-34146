VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save MultiGradient"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   4050
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2745
      TabIndex        =   5
      Top             =   1080
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK ! Save..."
      Height          =   330
      Left            =   2745
      TabIndex        =   4
      Top             =   675
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   90
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   270
      Width           =   3885
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gradients"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3615
      Left            =   90
      TabIndex        =   0
      Top             =   585
      Width           =   2445
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
         Height          =   3210
         Left            =   45
         Pattern         =   "*.mgr"
         TabIndex        =   1
         Top             =   315
         Width           =   2310
      End
   End
   Begin VB.Label Label1 
      Caption         =   "GradientName:"
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
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   45
      Width           =   1770
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() 'save gradient
On Error GoTo SaveError
MGFile = Text1
If LCase(Right(MGFile, 4)) <> ".mgr" Then MGFile = MGFile + ".Mgr"
For xx = 0 To File1.ListCount - 1
    If LCase(MGFile) = LCase(File1.List(xx)) Then
    Temp = MsgBox("The file " & MGFile & " allready exists." & vbCr & "Do you want to replace the file ?", vbYesNo + vbQuestion, MGTitle)
        If Temp = vbNo Then
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1)
        Text1.SetFocus
        Exit Sub
        End If
    End If
Next xx
For xx = 0 To 9
Mgpct2(xx) = MGPct(xx) * 10000
Next xx
ff = FreeFile
Open File1.Path & "\" & MGFile For Output As #ff
For xx = 0 To 9
Print #ff, MGrad(xx)
Print #ff, Mgpct2(xx)
Next xx
Close #ff
Form1.File1.Refresh
File1.Refresh
Form1.Label7 = "Last action: Save gradient"
Form2.Hide
Exit Sub
SaveError:
Close #ff
MsgBox "Error " & Err.Number & vbCr & Err.Description, , MGTitle
End Sub

Private Sub Command2_Click()
Form2.Hide
End Sub

Private Sub File1_Click()
Text1 = File1.List(File1.ListIndex)
End Sub

Private Sub Form_Activate()
File1.Refresh
Text1 = Form1.Label6
Text1.SelStart = 0
Text1.SelLength = Len(Text1)
Text1.SetFocus
End Sub

Private Sub Form_Load()
File1.Path = App.Path + "\Gradients"
End Sub
