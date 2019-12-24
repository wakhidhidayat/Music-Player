VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   " Music Player"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4500
   BeginProperty Font 
      Name            =   "MV Boli"
      Size            =   23.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3840
      Top             =   360
   End
   Begin VB.Frame Frame1 
      Caption         =   "Now Playing"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   3495
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   360
      TabIndex        =   4
      Top             =   5160
      Width           =   3615
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   3615
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   360
      TabIndex        =   2
      Top             =   3000
      Width           =   3615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3840
      Top             =   1320
   End
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      _Version        =   327682
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
File1.Path = Dir1
End Sub

Private Sub Drive1_Change()
Dir1 = Drive1
End Sub

Private Sub File1_DblClick()
Label1.Caption = File1.FileName
MMControl1.Command = "close"
MMControl1.FileName = File1.Path + "\" + File1.FileName
MMControl1.Command = "open"
MMControl1.Command = "play"
End Sub

Private Sub Form_Load()
MMControl1.BackVisible = False
MMControl1.StepVisible = False
MMControl1.RecordVisible = False
MMControl1.EjectVisible = False
File1.Pattern = "*.mp3"
End Sub

Private Sub MMControl1_Done(NotifyCode As Integer)
Dim x As Integer
On Error GoTo salah
If MMControl1.Position = MMControl1.Length Then
    x = File1.Selected(x) = True
    File1_DblClick
salah:
    Exit Sub
End If
End Sub
Private Sub Timer1_Timer()
If MMControl1.FileName <> "" Then
    Slider1.Max = MMControl1.Length
    Slider1.Value = MMControl1.Position
End If
End Sub

Private Sub Timer2_Timer()
If Label1.Caption <> "" Then
    judul = Label1.Caption
    panjang = Len(judul)
    judul = Right(judul, panjang - 1) & Left(judul, 1)
    Label1 = judul
    Label1.Refresh
End If
End Sub
