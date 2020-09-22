VERSION 5.00
Begin VB.Form frmTrainer 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "StarCraft: Brood War v1.06 Trainer - Using TMK engine."
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmSample.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmSample.frx":030A
   MousePointer    =   99  'Custom
   Picture         =   "frmSample.frx":0FD4
   ScaleHeight     =   4515
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton SendSCCommand 
      Caption         =   "Send Command to StarCraft"
      Height          =   375
      Left            =   2460
      TabIndex        =   9
      Top             =   3360
      Width           =   4095
   End
   Begin VB.Timer tmrBeatTS 
      Interval        =   100
      Left            =   1800
      Top             =   3240
   End
   Begin TrainerMakerKit.TrainerMaker TM 
      Left            =   6600
      Top             =   2280
      _ExtentX        =   1693
      _ExtentY        =   1693
   End
   Begin VB.Timer tmrAnimation 
      Interval        =   100
      Left            =   120
      Top             =   3240
   End
   Begin VB.Timer tmrHotKeys 
      Interval        =   1
      Left            =   1380
      Top             =   3240
   End
   Begin VB.CommandButton cmdTrainerOptions 
      Caption         =   "Tra&iner Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5220
      TabIndex        =   8
      Top             =   2940
      Width           =   1335
   End
   Begin VB.Timer tmrLockValues 
      Interval        =   500
      Left            =   960
      Top             =   3240
   End
   Begin VB.Timer tmrStatus 
      Interval        =   500
      Left            =   540
      Top             =   3240
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "A&bout"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   2940
      Width           =   1335
   End
   Begin VB.CheckBox cmdLockGas 
      Caption         =   "Lock"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5220
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CheckBox cmdLockMinerals 
      Caption         =   "Lock"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5220
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2100
      Width           =   1335
   End
   Begin VB.CommandButton RunSC 
      Caption         =   "R&un StarCraft"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2460
      TabIndex        =   6
      Top             =   2940
      Width           =   1335
   End
   Begin VB.CommandButton cmdModifyGas 
      Caption         =   "Modify"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdModifyMinerals 
      Caption         =   "Modify"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   2100
      Width           =   1335
   End
   Begin VB.CommandButton cmdGetGas 
      Caption         =   "Get Value"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2460
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdGetMinerals 
      Caption         =   "Get Value"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2460
      TabIndex        =   0
      Top             =   2100
      Width           =   1335
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Waiting for StarCraft"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3900
      TabIndex        =   16
      Top             =   3840
      Width           =   2715
   End
   Begin VB.Image SC_Icon 
      Height          =   480
      Left            =   2520
      Picture         =   "frmSample.frx":3B996
      Top             =   3720
      Width           =   480
   End
   Begin VB.Image bt_logo 
      Height          =   2475
      Left            =   7080
      MouseIcon       =   "frmSample.frx":3BCA0
      MousePointer    =   99  'Custom
      Picture         =   "frmSample.frx":3C56A
      ToolTipText     =   "Black Tornado Logo"
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label lblHotKey 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hotkey: F12"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   15
      Top             =   2880
      Width           =   870
   End
   Begin VB.Label lblHotKey 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hotkey: F11"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   14
      Top             =   2400
      Width           =   870
   End
   Begin VB.Image Image2 
      Height          =   210
      Left            =   120
      Picture         =   "frmSample.frx":3DBB6
      Top             =   2640
      Width           =   210
   End
   Begin VB.Image Image1 
      Height          =   210
      Left            =   120
      Picture         =   "frmSample.frx":3DE60
      Top             =   2160
      Width           =   210
   End
   Begin VB.Label lblGas 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   240
      Left            =   1560
      TabIndex        =   13
      Top             =   2640
      Width           =   630
   End
   Begin VB.Label lblMinerals 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   240
      Left            =   1560
      TabIndex        =   12
      Top             =   2160
      Width           =   630
   End
   Begin VB.Label lblInfo2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gas:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   480
      TabIndex        =   11
      Top             =   2640
      Width           =   390
   End
   Begin VB.Label lblInfo1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Minerals:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   480
      TabIndex        =   10
      Top             =   2160
      Width           =   795
   End
End
Attribute VB_Name = "frmTrainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Simple Trainer for Blizzard StarCraft
  ' Brood War v1.06 (Many not work with other versions)
  ' this template will browse TMK features for you, so enjoy!
  ' *************************************************************
  ' Some words on TMK
  ' -----------------
  ' TMK was made by Black Tornado, if you like it then vote for
  ' TMK on PSC, if you have any comments post on PSC or mail me
  ' at my e-mail address: btsoft@burntmail.com
  ' I have also created a forum, it is still in the begining but
  ' If you joined us you will be happy for that. Just visit the
  ' forum at : www.BlackTornado.cjb.net and join us and post your
  ' comments or questions about VB.
  ' I hope that you saw this project usefull, happy programming!
  ' Best Wishes
  '                       Black Tornado
  '           for me (NOTHING IS IMPOSSIBLE)
  ' *************************************************************
' Let us begin coding...
' Some important informations about StarCraft
Const GameCaption = "Brood War" ' StarCraft caption
Const GameClass = "SWarClass" ' StarCraft class
Const Minerals_Pos = &H4FEE58   ' Memory address of Minerals
Const Gas_Pos = &H4FEE88        ' Memory address of Gas
Dim Lock_Minerals As Boolean
Dim Lock_Gas As Boolean
Dim StarCraftRunning As Boolean
Dim AnimationKey As Long      ' Animation key saves frame number!!!
Dim AnimationDirection As String

Private Sub bt_logo_Click()
TM.About
End Sub

Private Sub cmdAbout_Click()
TM.About
End Sub

Private Sub cmdGetGas_Click()
lblGas.Caption = TM.ReadFromMemory(GameClass, Gas_Pos, 4)
End Sub

Private Sub cmdGetMinerals_Click()
lblMinerals.Caption = TM.ReadFromMemory(GameClass, Minerals_Pos, 4)
End Sub

Private Sub cmdLockGas_Click()
Lock_Gas = cmdLockGas.Value
End Sub

Private Sub cmdLockMinerals_Click()
Lock_Minerals = cmdLockMinerals.Value
End Sub

Private Sub cmdModifyGas_Click()
On Error GoTo Error
Lock_Gas = False
Dim New_GasValue As Long
New_GasValue = 0
New_GasValue = InputBox("Enter the new 'Gas' value:")
TM.WriteToMemory Gas_Pos, 4, New_GasValue, GameClass
Exit Sub
Error:
If StarCraftRunning = True Then MsgBox "Invalid Value!", vbExclamation: Exit Sub
MsgBox "StarCraft is not running!", vbCritical
End Sub

Private Sub cmdModifyMinerals_Click()
On Error GoTo Error
Lock_Minerals = False
Dim New_MineralsValue As Long
New_MineralsValue = 0
New_MineralsValue = InputBox("Enter the new 'Minerals' value:")
TM.WriteToMemory Minerals_Pos, 4, New_MineralsValue, GameClass
Exit Sub
Error:
If StarCraftRunning = True Then MsgBox "Invalid Value!", vbExclamation: Exit Sub
MsgBox "StarCraft is not running!", vbCritical
End Sub

Private Sub cmdTrainerOptions_Click()
On Error Resume Next
Dim NewInterval As Long
NewInterval = InputBox("Enter new 'Lock' interval:", "Set new timer interval", tmrLockValues.Interval)
tmrLockValues.Interval = NewInterval
End Sub

Private Sub Form_Load()
' Check if the application has ran more than once
If App.PrevInstance = True Then MsgBox "You can't run the trainer more than once at the same time", vbCritical: End
Screen.MouseIcon = Me.MouseIcon
Screen.MousePointer = Me.MousePointer
Lock_Minerals = False
Lock_Gas = False
AnimationKey = 2460
AnimationDirection = "LEFT"
If TM.CheckForTS(NewVersion) = True Then GoTo TS_Detected
If TM.CheckForTS(OldVersion) = True Then GoTo TS_Detected
ShowStarCraftStatus
If StarCraftRunning Then cmdGetMinerals_Click: cmdGetGas_Click
TM.TerminateTS ' Terminate even if TS was not detected, this protection is GOOD!!!
Exit Sub
TS_Detected: ' We saw 'Trainer Spy'
MsgBox "Why are you spying on me?", vbQuestion, "Trainer Maker Kit"
TM.TerminateTS
End Sub

Sub ShowStarCraftStatus()
lblStatus.Caption = IIf(TM.IsProgramRunning(GameCaption, GameClass, FindByClassAndCaption) = True, "StarCraft is running", "Waiting for StarCraft")
StarCraftRunning = TM.IsProgramRunning(GameCaption, GameClass, FindByClassAndCaption)
End Sub

Private Sub RunSC_Click()
TM.ExecuteProgram App.Path & "\StarCraft.exe", BT_SHOW, App.Path
End Sub

Private Sub SendSCCommand_Click()
MsgBox "Now I will send a command that will Disable/Enable GOD mode"
TM.ActivateWindow GameCaption, GameClass, bt_shownormal
TM.SendKeysToGame "{ENTER}"            ' when we will activate SC we must exit main menu
TM.SendKeysToGame "{ENTER}"            ' press enter so that starcraft says 'Message:'
TM.SendKeysToGame "power overwhelming" ' This is the cheat
TM.SendKeysToGame "{ENTER}"            ' execute it!
End Sub

Private Sub tmrAnimation_Timer()
If AnimationKey <= 2460 Then AnimationDirection = "RIGHT"
If AnimationKey >= 3300 Then AnimationDirection = "LEFT"
Select Case AnimationDirection:
Case "LEFT"
AnimationKey = AnimationKey - 50
Case "RIGHT"
AnimationKey = AnimationKey + 50
End Select
SC_Icon.Left = AnimationKey
End Sub

Private Sub tmrBeatTS_Timer()
' Always Kill TS
' This protection is good because the user may not run TS at the first time
' he may run TS after a while of running your Trainer
TM.TerminateTS
End Sub

Private Sub tmrHotKeys_Timer()
If StarCraftRunning = False Then Exit Sub
If TM.WaitHotKey(vbKeyF11) = True Then TM.WriteToMemory Minerals_Pos, 4, 999999, GameClass
If TM.WaitHotKey(vbKeyF12) = True Then TM.WriteToMemory Gas_Pos, 4, 999999, GameClass
End Sub

Private Sub tmrLockValues_Timer()
On Error Resume Next
If StarCraftRunning = False Then Exit Sub
If Lock_Minerals = True Then TM.WriteToMemory Minerals_Pos, 4, 999999, GameClass
If Lock_Gas = True Then TM.WriteToMemory Gas_Pos, 4, 999999, GameClass
End Sub

Private Sub tmrStatus_Timer()
Call ShowStarCraftStatus
End Sub
