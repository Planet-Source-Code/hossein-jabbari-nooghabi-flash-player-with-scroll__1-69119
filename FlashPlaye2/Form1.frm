VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flash Player"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6720
      Top             =   6360
   End
   Begin MSComctlLib.Slider SL 
      Height          =   675
      Left            =   1320
      TabIndex        =   6
      Top             =   6240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1191
      _Version        =   393216
      TextPosition    =   1
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   3540
      Left            =   7320
      Pattern         =   "*.swf"
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7320
      TabIndex        =   2
      Top             =   7800
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   3690
      Left            =   7320
      TabIndex        =   1
      Top             =   4080
      Width           =   2775
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash SF1 
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6855
      _cx             =   12091
      _cy             =   9763
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   "000000"
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Left            =   6000
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WEB SITE"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6120
      TabIndex        =   12
      Top             =   7080
      Width           =   705
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Right Reserved By HJN System                 Mobile: 09329289726"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   375
      TabIndex        =   11
      Top             =   7080
      Width           =   4755
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Right Reserved By HJN System                 Mobile: 09329289726"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   10
      Top             =   7065
      Width           =   4755
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "   "
      Height          =   195
      Left            =   2280
      TabIndex        =   9
      Top             =   7800
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Path Of File that Loaded:"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   7800
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "   "
      Height          =   195
      Left            =   2280
      TabIndex        =   7
      Top             =   7440
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   330
      Left            =   960
      Picture         =   "Form1.frx":5C12
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   285
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   240
      X2              =   6960
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   480
      Picture         =   "Form1.frx":5E74
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name Of File that Loaded:"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   7440
      Width           =   1890
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select File From Here:"
      Height          =   195
      Left            =   7320
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   5775
      Left            =   120
      Top             =   240
      Width           =   7095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   240
      X2              =   6960
      Y1              =   7320
      Y2              =   7320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DriveEr, DriveEr2 As String
Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
DriveEr = Drive1.Drive
On Error GoTo Errr:

Dir1.Path = Drive1.Drive
kk = UCase(Drive1.Drive)
Drive1drive = kk
DriveEr2 = Drive1.Drive

Exit Sub
Errr:
Call Drive3_Error
End Sub
Private Sub Drive3_Error()
  MsgBox "Error in load Drive " & DriveEr
  Drive1.Drive = DriveEr2

End Sub

Private Sub File1_Click()
SF1.Movie = File1.Path & "\" & File1.FileName
Label3.Caption = File1.FileName
Label4.Caption = File1.Path
SL.Max = SF1.TotalFrames

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape1.Visible = False
End Sub

Private Sub Image2_Click()
SF1.Play
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.BorderStyle = 1

End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.BorderStyle = 0
End Sub

Private Sub Image3_Click()
SF1.Stop
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.BorderStyle = 1
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.BorderStyle = 0
End Sub

Private Sub Label6_Click()
    Shell "explorer http://hosseinjn.googlepages.com/home"
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape1.Visible = True
End Sub

Private Sub SL_Click()

SF1.GotoFrame SL.Value
SF1.Play
End Sub

Private Sub SL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Interval = 0

End Sub

Private Sub SL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Interval = 1

End Sub

Private Sub SL_Scroll()
SF1.GotoFrame SL.Value
SF1.Play

End Sub

Private Sub Timer1_Timer()
'On Error Resume Next
SL.Value = SF1.CurrentFrame

End Sub
