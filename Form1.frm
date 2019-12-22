VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   " "
   ClientHeight    =   4155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   3360
      Top             =   1800
   End
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2520
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393216
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      _Version        =   327682
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CommonDialog1.ShowOpen
MMControl1.FileName = CommonDialog1.FileName
MMControl1.Command = "Open"
MMControl1.Command = "Play"
Command2.Caption = "Stop"
End Sub

Private Sub Command2_Click()
If Command1.Caption = "Play" Then
    MMControl1.Command = "Play"
    Command2.Caption = "Stop"
Else
    MMControl1.Command = "Stop"
    Command2.Caption = "Play"
End If
End Sub
