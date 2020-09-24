VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_main 
   Caption         =   "Transparent Subtitles over the Movie :) by : Üllar Luik"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10230
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9600
      Top             =   3600
   End
   Begin VB.CommandButton Command6 
      Caption         =   "pause"
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "stop"
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "play"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Width           =   1095
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   4080
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   661
      _Version        =   393216
      TickStyle       =   3
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9120
      Top             =   3600
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   8160
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Subtitle font settings"
      Height          =   375
      Left            =   8160
      TabIndex        =   4
      Top             =   1560
      Width           =   1935
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   3836
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Show"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Hide"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Text"
         Object.Width           =   13229
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open subtitle (mDVD)"
      Height          =   375
      Left            =   8160
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open movie file"
      Height          =   375
      Left            =   8160
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   4680
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frm_main.frx":1ADA
      Height          =   855
      Left            =   8160
      TabIndex        =   5
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000010&
      X1              =   8160
      X2              =   10080
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000014&
      X1              =   8160
      X2              =   10080
      Y1              =   2175
      Y2              =   2175
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      X1              =   8160
      X2              =   10080
      Y1              =   135
      Y2              =   135
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   8160
      X2              =   10080
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   8160
      X2              =   10080
      Y1              =   1335
      Y2              =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   8160
      X2              =   10080
      Y1              =   1320
      Y2              =   1320
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   0   'False
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   1
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -780
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim frame As Long

Private Sub Command1_Click()
    
    'open movie file
    
    With cd1
        
        .DialogTitle = "Open"
        .Filter = "Avi|*.avi"
        .Filename = ""
        .ShowOpen
        
        If .Filename = "" Then
            Exit Sub
        Else
            MediaPlayer1.Filename = .Filename
            Slider1.Max = MediaPlayer1.Duration
        End If
        
    End With
    
End Sub

Private Sub Command2_Click()

    'open subtitle file
    
    With cd1
        
        .DialogTitle = "Open"
        .Filter = "All types (*.*)|*.*"
        .Filename = ""
        .ShowOpen
        
        If .Filename = "" Then
            Exit Sub
        Else
            ListView1.ListItems.Clear
            Call OpenSub_mDVD(.Filename)
            subnr = 1
            frm_subtitle.Show
            frm_subtitle.Top = Me.Top + 3350
            frm_subtitle.Left = Me.Left + 160
        End If
        
    End With

End Sub

Private Sub Command4_Click()

'    On Error Resume Next
    
    MediaPlayer1.Play
    Timer1.Enabled = True
    Timer2.Enabled = True
    Read_mDVD_subtitle

End Sub

Private Sub Command5_Click()

    MediaPlayer1.Stop
    MediaPlayer1.CurrentPosition = 0
    Slider1.Value = 0
    subnr = 1

End Sub

Private Sub Command6_Click()

    MediaPlayer1.Pause

End Sub

Private Sub Form_Unload(Cancel As Integer)

    End

End Sub

Private Sub Slider1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Timer2.Enabled = False

End Sub

Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Timer2.Enabled = True
    MediaPlayer1.CurrentPosition = Slider1.Value
    DoEvents
    subnr = 1
    Read_mDVD_subtitle

End Sub

Private Sub Timer1_Timer()
    
    frame = MediaPlayer1.CurrentPosition
    Label2.Caption = frame
    
    If frame > startshow And frame < stopshow Then
        frm_subtitle.Label1.Caption = subtext
    End If
    
    If frame > stopshow Then
        frm_subtitle.Label1.Caption = ""
        Read_mDVD_subtitle
    End If
    
End Sub

Private Sub Timer2_Timer()

    Slider1.Value = MediaPlayer1.CurrentPosition

End Sub
