VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   " "
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   8040
      TabIndex        =   13
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Timer Timer5 
      Left            =   1440
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   495
      Left            =   7800
      TabIndex        =   12
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Repeat"
      Height          =   495
      Left            =   9000
      TabIndex        =   11
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer Timer4 
      Left            =   1080
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Left            =   720
      Top             =   0
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   7800
      TabIndex        =   10
      Top             =   1320
      Width           =   2295
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   1815
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   1815
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   2160
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1085
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2040
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   3240
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393216
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      BackEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label Label6 
      Caption         =   "00:00:00"
      Height          =   495
      Left            =   6720
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "dari"
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "JUDUL"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim File As String
Dim Kode As Boolean
Dim Endtrack As Long
Dim Jam, menit, detik, mldetik As Integer

Sub Play()
mldetik = 0
detik = 0
menit = 0
Jam = 0

File = List2

If Mid(File, 3, 1) = "\" And Mid(File, 4, 1) = "\" Then
    File = List1
Else
    File = List2
End If
MMControl1.FileName = File
MMControl1.Command = "Open"
Endtrack = MMControl1.TrackLength

If Endtrack = 0 Then
    MsgBox "Tidak dapat memutar file", vbOKOnly + vbCritical, "Player Error"
End If
End Sub

Private Sub Command1_Click()
List1.Clear
List2.Clear
End Sub

Private Sub Command2_Click()
If Command2.Caption = "Repeat" Then
    Command2.Caption = "Off"
Else
    Command2.Caption = "Repeat"
End If
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
File1.FileName = "*.mp3"
End Sub

Private Sub Drive1_Change()
On Error GoTo Perangkap
Dir1.Path = Drive1.Drive
Perangkap:
    Select Case Err
        Case 68
            MsgBox "Tidak dapat mengakses drive", vbOKOnly + vbCritical, "Scope Error"
            Drive1.Refresh
        Case 0
        Exit Sub
    End Select
End Sub

Private Sub File1_DblClick()
If File1.Name = "" Then
    Exit Sub
Else
    List1.AddItem File1.FileName
    List2.AddItem File1.Path & "\" & File1.FileName
    Label3.Caption = List1.ListIndex + 1
    Label5.Caption = List1.ListCount
End If
End Sub

Private Sub List1_Click()
List2.ListIndex = List1.ListIndex
Label3.Caption = List1.ListIndex + 1
Label5.Caption = List1.ListCount
Label2.Caption = List1
MMControl1.Command = "Close"
MMControl1.Refresh
End Sub

Private Sub MMControl1_Done(NotifyCode As Integer)
If Kode = True Then Exit Sub
If MMControl1.TrackLength = MMControl1.Position Then
    If Label3.Caption = Label5.Caption Then
        If Command2.Caption = "Repeat" Then
            If Label5.Caption = "1" Then
                MMControl1.Command = "Close"
                Timer2.Enabled = False
            Else
            If Label3.Caption = Label5.Caption Then
                List1.ListIndex = 0
                MMControl1.Command = "Play"
            End If
            End If
        Else
        If Label3.Caption = Label5.Caption Then
            MMControl1.Command = "Close"
        End If
        End If
    Else
    With List1.ListIndex = .ListIndex + 1
    End With
    MMControl1.Command = "Play"
End If
End If

End Sub

Private Sub MMControl1_PlayClick(Cancel As Integer)
Play
ProgressOke
Label2.Caption = List1
Timer2.Enabled = True
End Sub

Private Sub MMControl1_StopClick(Cancel As Integer)
MMControl1.Refresh
MMControl1.Command = "Close"
Kode = True
Timer2.Enabled = False
End Sub

Sub ProgressOke()
Slider1.Min = 0
Slider1.Max = Val(MMControl1.TrackLength)
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
Slider1.Value = MMControl1.Position
End Sub

Private Sub Timer2_Timer()
If mldetik = 10 Then
    detik = detik + 1
    mldetik = 0
End If
If detik = 60 Then
    menit = menit + 1
    detik = 0
End If
If menit = 60 Then
    Jam = Jam + 1
    menit = 0
End If
Label6.Caption = Jam & ":" & menit & ":" & detik
mldetik = mldetik + 1
End Sub

Private Sub Timer3_Timer()
If Me.Top <= 1000 Then
    Timer3.Interval = 0
Else
    Me.Top = Me.Top - 100
End If


End Sub

