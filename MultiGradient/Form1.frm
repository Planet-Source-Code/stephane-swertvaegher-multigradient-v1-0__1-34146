VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  MultiGradient V1.0 - ©2002 by Stephan Swertvaegher"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   429
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   525
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   135
      TabIndex        =   40
      Top             =   1935
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   23
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyNew"
            Object.ToolTipText     =   "Make new gradient (reset)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keySave"
            Object.ToolTipText     =   "Save gradient"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyKill"
            Object.ToolTipText     =   "Delete gradient from file"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyRename"
            Object.ToolTipText     =   "Rename gradient"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyVB"
            Object.ToolTipText     =   "Copy code"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyShift"
            Object.ToolTipText     =   "Shift colors"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyReverse"
            Object.ToolTipText     =   "Reverse colors"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyNeg"
            Object.ToolTipText     =   "Negative colors"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyRed"
            Object.ToolTipText     =   "Manipulate red component"
            ImageIndex      =   13
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "keyKillRed"
                  Text            =   "Kill Red"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "keyHalfRed"
                  Text            =   "Half Red"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "keyDoubleRed"
                  Text            =   "Double Red"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyGreen"
            Object.ToolTipText     =   "Manipulate green component"
            ImageIndex      =   14
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "keyKillGreen"
                  Text            =   "Kill Green"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "keyHalfGreen"
                  Text            =   "Half Green"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "keyDoubleGreen"
                  Text            =   "Double Green"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyBlue"
            Object.ToolTipText     =   "Manipulate blue component"
            ImageIndex      =   15
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "keyKillBlue"
                  Text            =   "Kill Blue"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "keyHalfBlue"
                  Text            =   "Half Blue"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "keyDoubleBlue"
                  Text            =   "Double Blue"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyRGB"
            Object.ToolTipText     =   "Manipulate RGB"
            ImageIndex      =   16
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "keyRGBRBG"
                  Text            =   "RGB ---> RBG"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "keyRGBGBR"
                  Text            =   "RGB ---> GBR"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "keyRGBGRB"
                  Text            =   "RGB ---> GRB"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "keyRGBBGR"
                  Text            =   "RGB ---> BGR"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "keyRGBBRG"
                  Text            =   "RGB ---> BRG"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keySelect"
            Object.ToolTipText     =   "Select all colors"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyDeselect"
            Object.ToolTipText     =   "Deselect all colors"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyFull"
            Object.ToolTipText     =   "View gradient"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyHelp"
            Object.ToolTipText     =   "Show help"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   810
      Top             =   495
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0466
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":05C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0A16
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0B72
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0CCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0E2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1C7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1DDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1F36
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2092
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":21EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":234A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":24A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2602
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":275E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
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
      Height          =   3345
      Left            =   4140
      TabIndex        =   38
      Top             =   2520
      Width           =   3660
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
         Height          =   3015
         Left            =   225
         Pattern         =   "*.mgr"
         TabIndex        =   39
         Top             =   270
         Width           =   3210
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pointer Info"
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
      Height          =   3345
      Left            =   135
      TabIndex        =   4
      Top             =   2520
      Width           =   3975
      Begin VB.CheckBox Check1 
         Caption         =   "Fix"
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
         Index           =   9
         Left            =   3240
         TabIndex        =   53
         Top             =   3060
         Value           =   1  'Checked
         Width           =   645
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Index           =   8
         Left            =   3240
         TabIndex        =   52
         Top             =   2790
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Index           =   7
         Left            =   3240
         TabIndex        =   51
         Top             =   2520
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Index           =   6
         Left            =   3240
         TabIndex        =   50
         Top             =   2250
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Index           =   5
         Left            =   3240
         TabIndex        =   49
         Top             =   1980
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Index           =   4
         Left            =   3240
         TabIndex        =   48
         Top             =   1710
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Index           =   3
         Left            =   3240
         TabIndex        =   47
         Top             =   1440
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Index           =   2
         Left            =   3240
         TabIndex        =   46
         Top             =   1170
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Index           =   1
         Left            =   3240
         TabIndex        =   45
         Top             =   900
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Fix"
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
         Index           =   0
         Left            =   3240
         TabIndex        =   44
         Top             =   630
         Value           =   1  'Checked
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Enabled"
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
         Height          =   240
         Index           =   3
         Left            =   3015
         TabIndex        =   43
         Top             =   315
         Width           =   690
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   90
         Picture         =   "Form1.frx":28BA
         Top             =   585
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Slider"
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
         Height          =   240
         Index           =   0
         Left            =   360
         TabIndex        =   37
         Top             =   315
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Backcolor"
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
         Height          =   240
         Index           =   1
         Left            =   1080
         TabIndex        =   36
         Top             =   315
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Position"
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
         Height          =   240
         Index           =   2
         Left            =   2205
         TabIndex        =   35
         Top             =   315
         Width           =   690
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Height          =   240
         Index           =   0
         Left            =   405
         TabIndex        =   34
         Top             =   585
         Width           =   465
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   0
         Left            =   1275
         TabIndex        =   33
         Top             =   585
         Width           =   555
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   240
         Index           =   0
         Left            =   2160
         TabIndex        =   32
         Top             =   585
         Width           =   780
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
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
         Height          =   240
         Index           =   1
         Left            =   2160
         TabIndex        =   31
         Top             =   855
         Width           =   780
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   1
         Left            =   1275
         TabIndex        =   30
         Top             =   855
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
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
         Height          =   240
         Index           =   1
         Left            =   405
         TabIndex        =   29
         Top             =   855
         Width           =   465
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
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
         Height          =   240
         Index           =   2
         Left            =   2160
         TabIndex        =   28
         Top             =   1125
         Width           =   780
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   2
         Left            =   1275
         TabIndex        =   27
         Top             =   1125
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
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
         Height          =   240
         Index           =   2
         Left            =   405
         TabIndex        =   26
         Top             =   1125
         Width           =   465
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
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
         Height          =   240
         Index           =   3
         Left            =   2160
         TabIndex        =   25
         Top             =   1395
         Width           =   780
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   3
         Left            =   1275
         TabIndex        =   24
         Top             =   1395
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
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
         Height          =   240
         Index           =   3
         Left            =   405
         TabIndex        =   23
         Top             =   1395
         Width           =   465
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
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
         Height          =   240
         Index           =   4
         Left            =   2160
         TabIndex        =   22
         Top             =   1665
         Width           =   780
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   4
         Left            =   1275
         TabIndex        =   21
         Top             =   1665
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
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
         Height          =   240
         Index           =   4
         Left            =   405
         TabIndex        =   20
         Top             =   1665
         Width           =   465
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
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
         Height          =   240
         Index           =   5
         Left            =   2160
         TabIndex        =   19
         Top             =   1935
         Width           =   780
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   5
         Left            =   1275
         TabIndex        =   18
         Top             =   1935
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
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
         Height          =   240
         Index           =   5
         Left            =   405
         TabIndex        =   17
         Top             =   1935
         Width           =   465
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
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
         Height          =   240
         Index           =   6
         Left            =   2160
         TabIndex        =   16
         Top             =   2205
         Width           =   780
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   6
         Left            =   1275
         TabIndex        =   15
         Top             =   2205
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
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
         Height          =   240
         Index           =   6
         Left            =   405
         TabIndex        =   14
         Top             =   2205
         Width           =   465
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
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
         Height          =   240
         Index           =   7
         Left            =   2160
         TabIndex        =   13
         Top             =   2475
         Width           =   780
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   7
         Left            =   1275
         TabIndex        =   12
         Top             =   2475
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "7"
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
         Height          =   240
         Index           =   7
         Left            =   405
         TabIndex        =   11
         Top             =   2475
         Width           =   465
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
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
         Height          =   240
         Index           =   8
         Left            =   2160
         TabIndex        =   10
         Top             =   2745
         Width           =   780
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   8
         Left            =   1275
         TabIndex        =   9
         Top             =   2745
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
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
         Height          =   240
         Index           =   8
         Left            =   405
         TabIndex        =   8
         Top             =   2745
         Width           =   465
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   240
         Index           =   9
         Left            =   2160
         TabIndex        =   7
         Top             =   3015
         Width           =   780
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   9
         Left            =   1275
         TabIndex        =   6
         Top             =   3015
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9"
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
         Height          =   240
         Index           =   9
         Left            =   405
         TabIndex        =   5
         Top             =   3015
         Width           =   465
      End
   End
   Begin VB.PictureBox Pic3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   180
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   2
      ToolTipText     =   "Move the sliders to have different positions of the gradient"
      Top             =   1530
      Width           =   7500
      Begin VB.PictureBox Sli 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   150
         Index           =   0
         Left            =   0
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   5
         TabIndex        =   3
         Top             =   0
         Width           =   75
      End
   End
   Begin VB.PictureBox Pic2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   180
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   1
      Top             =   1350
      Width           =   7500
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   180
      ScaleHeight     =   46
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   0
      ToolTipText     =   "Gradient"
      Top             =   540
      Width           =   7500
      Begin MSComDlg.CommonDialog CD1 
         Left            =   1800
         Top             =   45
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Last action:"
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
      Height          =   195
      Left            =   270
      TabIndex        =   54
      Top             =   6075
      Width           =   6495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   1710
      TabIndex        =   42
      Top             =   105
      Width           =   4740
   End
   Begin VB.Label Label5 
      Caption         =   "Gradient Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   225
      TabIndex        =   41
      Top             =   135
      Width           =   1365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Dim idx0%, Idx1%

Private Sub Check1_Click(Index As Integer)
Image1.Top = Label2(Index).Top + 1
    If Index = 0 Or Index = 9 Then
    Check1(0).Value = 1
    Check1(9).Value = 1
    Exit Sub
    End If
If Check1(Index).Value = 0 Then
Sli(Index).Visible = False
Label4(Index) = "***"
MGPct(Index) = -1
Else
Sli(Index).Visible = True
GetNextIndex Index
MGPos(Index) = (MGPos(idx0) + MGPos(Idx1)) / 2
MGPct(Index) = (MGPos(Index) + 2) / Pic1.ScaleWidth
Label4(Index) = Format(MGPct(Index), "0.0000")
Sli(Index).Left = MGPos(Index) - 2
End If
Pic1.Cls
If LoadFile = False Then MultiGrad2 Pic1
If Check1(Index).Value = 0 Then
Label7 = "Last action: Deselect slider " & Index
Else
Label7 = "Last action: Select slider " & Index
End If
If Start = True Then Label7 = "Welcome !"
End Sub

Private Sub File1_Click()
If KillFile = True Then
    KillFile = False
    MGFile = File1.List(File1.ListIndex)
    Temp = MsgBox("Do you want to delete the gradient " & File1.List(File1.ListIndex), vbOKCancel + vbQuestion, MGTitle)
    If Temp = vbCancel Then Exit Sub
    Kill File1.Path & "\" & MGFile
    Label7 = "Last action: Delete file " & MGFile
    File1.Refresh
    Exit Sub
End If
'-------------------------
LoadFile = True
Temp = MsgBox("Load the gradient: " & File1.List(File1.ListIndex), vbYesNo + vbQuestion, MGTitle)
If Temp = vbNo Then Exit Sub
On Error GoTo LoadError
    For xx = 0 To 9
    Check1(xx).Value = 1
    Sli(xx).Visible = True
    Next xx
ff = FreeFile
MGFile = File1.List(File1.ListIndex)
Open File1.Path & "\" & MGFile For Input As #ff
For xx = 0 To 9
Input #ff, MGrad(xx)
Label3(xx).BackColor = MGrad(xx)
Sli(xx).BackColor = MGrad(xx)
Input #ff, Mgpct2(xx)
MGPct(xx) = Mgpct2(xx) / 10000
    If MGPct(xx) = -1 Then
    Check1(xx).Value = 0
    Sli(xx).Visible = False
    End If
Next xx
For xx = 0 To 9
MGPos(xx) = (MGPct(xx) * Pic1.ScaleWidth) - 2
Sli(xx).Left = MGPos(xx)
Label4(xx) = Format(MGPct(xx), "0.0000")
If MGPct(xx) = -1 Then Label4(xx) = "***"
Next xx
Close #ff
LoadFile = False
MultiGrad2 Pic1
Image1.Top = Label2(0).Top + 1
Label6.Caption = MGFile
Label7 = "Last action: Load multigradient " & MGFile
Exit Sub
LoadError:
Close #ff
LoadFile = False
Image1.Top = Label2(0).Top + 1
MsgBox "Error " & Err.Number & vbCr & Err.Description, , MGTitle
End Sub

Private Sub Form_Activate()
Start = False
End Sub

Private Sub Form_Load()
MGTitle = "MultiGradient V1.0"
Start = True
Label7 = ""
Image1.Top = Label2(0).Top + 1
For xx = 0 To 49
Pic2.Line (xx * 10, 5)-(xx * 10, 10), &HA0
Next xx
For xx = 0 To 4
Pic2.Line (xx * 100, 0)-(xx * 100, 10), &HFF
Next xx
For xx = 1 To 9
Load Sli(xx)
Sli(xx).Visible = True
Sli(xx).ToolTipText = "Slider " & xx
Next xx
C1 = 200: C2 = 200: C3 = 222
For xx = 0 To 4
Pic3.Line (0, xx)-(499, xx), RGB(C1, C2, C3)
Pic3.Line (0, 9 - xx)-(499, 9 - xx), RGB(C1, C2, C3)
C1 = C1 - 10
C2 = C2 - 10
C3 = C3 - 10
Next xx
Form1.Line (0, 3)-(Form1.ScaleWidth, 3), &H808080
Form1.Line (0, 4)-(Form1.ScaleWidth, 4), &HE0E0E0
Form1.Line (0, 124)-(Form1.ScaleWidth, 124), &H808080
Form1.Line (0, 125)-(Form1.ScaleWidth, 125), &HE0E0E0
Form1.Line (0, 158)-(Form1.ScaleWidth, 158), &H808080
Form1.Line (0, 159)-(Form1.ScaleWidth, 159), &HE0E0E0
File1.Path = App.Path & "\Gradients"
NewGrad
MultiGrad2 Pic1
T3D Form1, Label7, 5, T3dRaiseInset
T3D Form1, Pic1, 5, T3dRaiseInset
Form1.Show
Form7.Show 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
Temp = MsgBox("Do you really want to leave ?", vbYesNo + vbQuestion, MGTitle)
If Temp = vbNo Then Exit Sub
End
End Sub

Private Sub Label2_Click(Index As Integer)
Image1.Top = Label2(Index).Top + 1
Label7 = "Last action: Select slider " & Index
End Sub

Private Sub Label3_Click(Index As Integer)
Image1.Top = Label2(Index).Top + 1
If Check1(Index).Value = 0 Then Exit Sub
CD1.Flags = 3
CD1.Color = Label3(Index).BackColor
CD1.ShowColor
Label3(Index).BackColor = CD1.Color
MGrad(Index) = CD1.Color
Sli(Index).BackColor = CD1.Color
Label7 = "Last action: change color of slider " & Index
Pic1.Cls
MultiGrad2 Pic1
End Sub

Private Sub Label4_Click(Index As Integer)
Image1.Top = Label2(Index).Top + 1
Label7 = "Last action: Select slider " & Index
End Sub

Private Sub Sli_Click(Index As Integer)
Image1.Top = Label2(Index).Top + 1
Label7 = "Last action: Select slider " & Index
End Sub

Private Sub Sli_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Top = Label2(Index).Top + 1
End Sub

Private Sub Sli_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Or Index = 9 Then Exit Sub
    Dim lngReturnValue As Long
    If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(Sli(Index).hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
Sli(Index).Top = 0
GetNextIndex Index
If Sli(Index).Left <= Sli(idx0).Left + 1 Then Sli(Index).Left = Sli(idx0).Left + 1
If Sli(Index).Left >= Sli(Idx1).Left - 1 Then Sli(Index).Left = Sli(Idx1).Left - 1
MGPos(Index) = Sli(Index).Left
MGPct(Index) = (MGPos(Index) + 2) / Pic1.ScaleWidth
Label4(Index) = Format(MGPct(Index), "0.0000")
DoEvents
Pic1.Cls
MultiGrad2 Pic1
Label7 = "Last action: Move slider " & Index
End If
End Sub

Private Sub GetNextIndex(idx%)
    'get first smaller pointer
    For qq = idx - 1 To 0 Step -1
    If Sli(qq).Visible = True Then
    idx0 = qq
    GoTo GetSecond
    Exit For
    End If
    Next qq
GetSecond:
    'get first bigger pointer
    For qq = idx + 1 To 9
    If Sli(qq).Visible = True Then
    Idx1 = qq
    GoTo Done
    End If
    Next qq
Done:
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "keyNew"
    Temp = MsgBox("Are you sure to begin a new gradient ?", vbYesNo + vbQuestion, MGTitle)
    If Temp = vbNo Then Exit Sub
    NewGrad
    MultiGrad2 Pic1
    Label7 = "Last action: Reset gradient (new gradient)"
Case "keySave"
    Form2.Show 1
Case "keyShift"
    TempCol = MGrad(9)
    For xx = 9 To 1 Step -1
    MGrad(xx) = MGrad(xx - 1)
    Next xx
    MGrad(0) = TempCol
    CopyColors
    Label7 = "Last action: Shift colors"
Case "keySelect"
    For xx = 0 To 9
    Check1(xx).Value = 1
    Next xx
    Label7 = "Last action: Select all sliders"
Case "keyDeselect"
For xx = 1 To 8
    Check1(xx).Value = 0
    Next xx
    Label7 = "Last action: Deselect all sliders"
Case "keyReverse"
    For xx = 0 To 4
    TempCol = MGrad(xx)
    MGrad(xx) = MGrad(9 - xx)
    MGrad(9 - xx) = TempCol
    Next xx
    CopyColors
    Label7 = "Last action: Reverse colors"
Case "keyNeg"
    For xx = 0 To 9
        If MGPct(xx) <> -1 Then
        R1 = MGrad(xx) Mod 256&
        G1 = ((MGrad(xx) And &HFF00) / 256&) Mod 256&
        B1 = (MGrad(xx) And &HFF0000) / 65536
        R1 = R1 Xor 255
        G1 = G1 Xor 255
        B1 = B1 Xor 255
        MGrad(xx) = RGB(R1, G1, B1)
        End If
    Next xx
    CopyColors
    Label7 = "Last action: Negative colors"
Case "keyKill"
    KillFile = True
    Temp = MsgBox("Select gradient in the list", vbInformation + vbOKCancel, MGTitle)
        If Temp = vbCancel Then
        KillFile = False
        Exit Sub
        End If
Case "keyRename"
    Form3.Show 1
Case "keyFull"
    MultiGrad Form4.Pic1, 0, False
    Form4.Show 1
Case "keyVB"
    Form5.Show 1
Case "keyHelp"
    Form6.Show 1
End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key
Case "keyKillRed"
    For xx = 0 To 9
        If MGPct(xx) <> -1 Then
        R1 = MGrad(xx) Mod 256&
        G1 = ((MGrad(xx) And &HFF00) / 256&) Mod 256&
        B1 = (MGrad(xx) And &HFF0000) / 65536
        R1 = 0
        MGrad(xx) = RGB(R1, G1, B1)
        End If
    Next xx
    Label7 = "Last action: kill red component"
Case "keyHalfRed"
    For xx = 0 To 9
        If MGPct(xx) <> -1 Then
        R1 = MGrad(xx) Mod 256&
        G1 = ((MGrad(xx) And &HFF00) / 256&) Mod 256&
        B1 = (MGrad(xx) And &HFF0000) / 65536
        R1 = R1 / 2
        MGrad(xx) = RGB(R1, G1, B1)
        End If
    Next xx
    Label7 = "Last action: half red component"
Case "keyDoubleRed"
    For xx = 0 To 9
        If MGPct(xx) <> -1 Then
        R1 = MGrad(xx) Mod 256&
        G1 = ((MGrad(xx) And &HFF00) / 256&) Mod 256&
        B1 = (MGrad(xx) And &HFF0000) / 65536
        R1 = R1 * 2
        If R1 > 255 Then R1 = 255
        MGrad(xx) = RGB(R1, G1, B1)
        End If
    Next xx
    Label7 = "Last action: double red component"
Case "keyKillGreen"
    For xx = 0 To 9
        If MGPct(xx) <> -1 Then
        R1 = MGrad(xx) Mod 256&
        G1 = ((MGrad(xx) And &HFF00) / 256&) Mod 256&
        B1 = (MGrad(xx) And &HFF0000) / 65536
        G1 = 0
        MGrad(xx) = RGB(R1, G1, B1)
        End If
    Next xx
    Label7 = "Last action: kill green component"
Case "keyHalfGreen"
    For xx = 0 To 9
        If MGPct(xx) <> -1 Then
        R1 = MGrad(xx) Mod 256&
        G1 = ((MGrad(xx) And &HFF00) / 256&) Mod 256&
        B1 = (MGrad(xx) And &HFF0000) / 65536
        G1 = G1 / 2
        MGrad(xx) = RGB(R1, G1, B1)
        End If
    Next xx
    Label7 = "Last action: half green component"
Case "keyDoubleGreen"
    For xx = 0 To 9
        If MGPct(xx) <> -1 Then
        R1 = MGrad(xx) Mod 256&
        G1 = ((MGrad(xx) And &HFF00) / 256&) Mod 256&
        B1 = (MGrad(xx) And &HFF0000) / 65536
        G1 = G1 * 2
        If G1 > 255 Then G1 = 255
        MGrad(xx) = RGB(R1, G1, B1)
        End If
    Next xx
    Label7 = "Last action: double green component"
Case "keyKillBlue"
    For xx = 0 To 9
        If MGPct(xx) <> -1 Then
        R1 = MGrad(xx) Mod 256&
        G1 = ((MGrad(xx) And &HFF00) / 256&) Mod 256&
        B1 = (MGrad(xx) And &HFF0000) / 65536
        B1 = 0
        MGrad(xx) = RGB(R1, G1, B1)
        End If
    Next xx
    Label7 = "Last action: kill blue component"
Case "keyHalfBlue"
    For xx = 0 To 9
        If MGPct(xx) <> -1 Then
        R1 = MGrad(xx) Mod 256&
        G1 = ((MGrad(xx) And &HFF00) / 256&) Mod 256&
        B1 = (MGrad(xx) And &HFF0000) / 65536
        B1 = B1 / 2
        MGrad(xx) = RGB(R1, G1, B1)
        End If
    Next xx
    Label7 = "Last action: half blue component"
Case "keyDoubleBlue"
    For xx = 0 To 9
        If MGPct(xx) <> -1 Then
        R1 = MGrad(xx) Mod 256&
        G1 = ((MGrad(xx) And &HFF00) / 256&) Mod 256&
        B1 = (MGrad(xx) And &HFF0000) / 65536
        B1 = B1 * 2
        If B1 > 255 Then B1 = 255
        MGrad(xx) = RGB(R1, G1, B1)
        End If
    Next xx
    Label7 = "Last action: double blue component"
Case "keyRGBRBG"
    For xx = 0 To 9
        If MGPct(xx) <> -1 Then
        R1 = MGrad(xx) Mod 256&
        G1 = ((MGrad(xx) And &HFF00) / 256&) Mod 256&
        B1 = (MGrad(xx) And &HFF0000) / 65536
        MGrad(xx) = RGB(R1, B1, G1)
        End If
    Next xx
    Label7 = "Last action: RGB to RBG"
Case "keyRGBGBR"
    For xx = 0 To 9
        If MGPct(xx) <> -1 Then
        R1 = MGrad(xx) Mod 256&
        G1 = ((MGrad(xx) And &HFF00) / 256&) Mod 256&
        B1 = (MGrad(xx) And &HFF0000) / 65536
        MGrad(xx) = RGB(G1, B1, R1)
        End If
    Next xx
    Label7 = "Last action: RGB to GBR"
Case "keyRGBGRB"
    For xx = 0 To 9
        If MGPct(xx) <> -1 Then
        R1 = MGrad(xx) Mod 256&
        G1 = ((MGrad(xx) And &HFF00) / 256&) Mod 256&
        B1 = (MGrad(xx) And &HFF0000) / 65536
        MGrad(xx) = RGB(G1, R1, B1)
        End If
    Next xx
    Label7 = "Last action: RGB to GRB"
Case "keyRGBBGR"
    For xx = 0 To 9
        If MGPct(xx) <> -1 Then
        R1 = MGrad(xx) Mod 256&
        G1 = ((MGrad(xx) And &HFF00) / 256&) Mod 256&
        B1 = (MGrad(xx) And &HFF0000) / 65536
        MGrad(xx) = RGB(B1, G1, R1)
        End If
    Next xx
    Label7 = "Last action: RGB to BGR"
Case "keyRGBBRG"
    For xx = 0 To 9
        If MGPct(xx) <> -1 Then
        R1 = MGrad(xx) Mod 256&
        G1 = ((MGrad(xx) And &HFF00) / 256&) Mod 256&
        B1 = (MGrad(xx) And &HFF0000) / 65536
        MGrad(xx) = RGB(B1, R1, G1)
        End If
    Next xx
    Label7 = "Last action: RGB to BRG"

End Select
    CopyColors
End Sub

Private Sub CopyColors()
    For xx = 0 To 9
    Label3(xx).BackColor = MGrad(xx)
    Sli(xx).BackColor = MGrad(xx)
    Next xx
    MultiGrad2 Pic1
End Sub
