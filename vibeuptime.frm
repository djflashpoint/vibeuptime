VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "V I B E U P T I M E"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   Icon            =   "vibeuptime.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Left            =   120
      Top             =   2520
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "MM/DD/YYYY HH:mm:ss"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   2535
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   6975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "00 Days 00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long


Private Sub Form_Load()
Label1 = GetSysUptime
Timer1.Interval = 100
Timer2.Interval = 100
Label1.ToolTipText = "System Uptime"
Label2.ToolTipText = "System Date/Time"
End Sub

Private Function GetSysUptime() As String
Dim lSec As Long
lSec = GetTickCount \ 1000
GetSysUptime = Format$((lSec \ 3600) \ 24, "00 Days" & vbNewLine) _
& Format$((lSec \ 3600) Mod 24, "00:") _
& Format$((lSec Mod 3600) \ 60, "00:") _
& Format$((lSec Mod 60), "00")
End Function

Private Sub Form_Resize()
Label1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
Label2.Move 0, 2400, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Label1_Click()
'
End Sub

Private Sub Label2_Click()
'
End Sub

Private Sub Timer1_Timer()
Label1 = GetSysUptime
End Sub

Private Sub Timer2_Timer()
Label2 = Now
Label2.Caption = Format(Now, "MM/DD/YYYY" & vbNewLine & "HH:mm:ss")
End Sub
