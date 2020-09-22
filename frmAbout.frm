VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Trainer Maker Kit"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8310
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   4020
      Width           =   7875
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "What is Trainer Spy?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   9
         Top             =   240
         Width           =   1740
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":237A
         ForeColor       =   &H00008080&
         Height          =   1275
         Index           =   6
         Left            =   180
         TabIndex        =   8
         Top             =   420
         Width           =   7470
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   3030
      Width           =   7875
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "These functions will help you very much building your trainer"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   4
         Left            =   3000
         TabIndex        =   6
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Detect Trainer Spy, Terminate Trainer Spy, Is Key Pressed, Terminate Window, Execute Program, Activate Window, Send Keys To Game."
         ForeColor       =   &H00008000&
         Height          =   435
         Index           =   3
         Left            =   180
         TabIndex        =   5
         Top             =   420
         Width           =   7470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Read From Memory, Write To Memory,"
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   2
         Left            =   4860
         TabIndex        =   4
         Top             =   240
         Width           =   2790
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trainer Maker Kit provides you many easy-to-use function like :"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   240
         Width           =   4545
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Index           =   4
      Left            =   240
      TabIndex        =   13
      Top             =   7320
      Width           =   7875
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Some words."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   15
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":25B5
         ForeColor       =   &H00008080&
         Height          =   855
         Index           =   9
         Left            =   180
         TabIndex        =   14
         Top             =   420
         Width           =   7470
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   7875
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":2748
         ForeColor       =   &H00808080&
         Height          =   1275
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   240
         Width           =   7515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   5730
      Width           =   7875
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":29A1
         ForeColor       =   &H00008080&
         Height          =   1155
         Index           =   7
         Left            =   180
         TabIndex        =   12
         Top             =   420
         Width           =   7470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Why did you Made 'Trainer Maker Kit'?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   11
         Top             =   240
         Width           =   3240
      End
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   990
      Left            =   2520
      Picture         =   "frmAbout.frx":2B94
      Top             =   240
      Width           =   3585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   11
      Left            =   2445
      TabIndex        =   19
      Top             =   8760
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Website:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   12
      Left            =   2280
      TabIndex        =   18
      Top             =   8940
      Width           =   735
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "btsoft@burntmail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   195
      Index           =   13
      Left            =   3480
      MousePointer    =   1  'Arrow
      TabIndex        =   17
      Top             =   8760
      Width           =   1890
   End
   Begin VB.Label lblWebsite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.BlackTornado.cjb.net"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   195
      Index           =   14
      Left            =   3120
      MousePointer    =   1  'Arrow
      TabIndex        =   16
      Top             =   8940
      Width           =   2820
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   960
      Left            =   240
      Picture         =   "frmAbout.frx":E576
      ToolTipText     =   "Click me to exit"
      Top             =   300
      Width           =   960
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub
Private Sub lblEmail_Click(Index As Integer)
ShellExecute Me.hwnd, "open", "mailto:btsoft@burntmail.com?Subject=Comments", 0, "C:\", 5
End Sub

Private Sub lblWebsite_Click(Index As Integer)
ShellExecute Me.hwnd, "open", "http://www.BlackTornado.cjb.net", 0, "C:\", 5
End Sub
