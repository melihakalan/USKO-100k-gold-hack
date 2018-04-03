VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.OCX"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Initializing..."
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3540
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   3540
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmLatency 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "KoJD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   2595
   End
   Begin VB.Label lblInit 
      Alignment       =   2  'Center
      Caption         =   "Init Message"
      Height          =   435
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2595
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
LoadMe
LoadProgram
End Sub

Private Sub Form_Load()
Load Form1
Form1.Hide
lblInit.Caption = "Initializing Program..." & vbCrLf & "Please Wait..."
End Sub

Public Sub LoadMe()

myVersion = "74510"
myName = Form1.Caption

err_Name = False
err_Run = False
err_Ver = False
err_Pass = False

End Sub

Public Sub LoadProgram()


ALLSETTING = Inet1.OpenURL("http://kojd.dabbledb.com/publish/100kdupe/bfc95bf8-adaf-464c-951d-259809379100/kojddupe.txt")

NameHead = InStr(1, ALLSETTING, "[hck", 0)
RunHead = InStr(1, ALLSETTING, "[run", 0)
VerHead = InStr(1, ALLSETTING, "[ver", 0)
NoRunMsgHead = InStr(1, ALLSETTING, "[msg", 0)
PassHead = InStr(1, ALLSETTING, "[pwd", 0)
WebOnExitHead = InStr(1, ALLSETTING, "[web", 0)
NoticeHead = InStr(1, ALLSETTING, "[ntc", 0)
RunWebHead = InStr(1, ALLSETTING, "[dwn", 0)

webName = Mid(ALLSETTING, NameHead + 8, val(Mid(ALLSETTING, NameHead + 5, 3)))
webRun = Mid(ALLSETTING, RunHead + 8, val(Mid(ALLSETTING, RunHead + 5, 3)))
webVersion = Mid(ALLSETTING, VerHead + 8, val(Mid(ALLSETTING, VerHead + 5, 3)))
webNORunMsg = Mid(ALLSETTING, NoRunMsgHead + 8, val(Mid(ALLSETTING, NoRunMsgHead + 5, 3)))
webPass = Mid(ALLSETTING, PassHead + 8, val(Mid(ALLSETTING, PassHead + 5, 3)))
webWebOnExit = Mid(ALLSETTING, WebOnExitHead + 8, val(Mid(ALLSETTING, WebOnExitHead + 5, 3)))
webNotice = Mid(ALLSETTING, NoticeHead + 8, val(Mid(ALLSETTING, NoticeHead + 5, 3)))
webRunWeb = Mid(ALLSETTING, RunWebHead + 8, val(Mid(ALLSETTING, RunWebHead + 5, 3)))

CheckMe

ResumeIfOK

Exit Sub
Error:
MsgBox "Initializing failed. Could not connect and get settings from web!", vbExclamation, "Error!"
End
End Sub

Public Sub CheckMe()

If myName <> webName Then
err_Name = True
GoTo errName
End If

If webRun = 0 Then
err_Run = True
GoTo errRun
End If

If myVersion <> webVersion Then
err_Ver = True
GoTo errVer
End If

If webPass <> 0 Then
err_Pass = True
CheckPWD
End If

Exit Sub

errName:
MsgBox "This hack is created by KoJD. It looks like edited. Please download original one." & vbCrLf & "~KoJD", vbExclamation, "EDIT!"
If webRunWeb <> 0 Then Shell "explorer " & webRunWeb
End

errRun:
MsgBox webNORunMsg & vbCrLf & "~KoJD", vbExclamation, "Not Available for now."
If webRunWeb <> 0 Then Shell "explorer " & webRunWeb
End

errVer:
MsgBox "Your version is outdated. Please download latest version of this hack." & vbCrLf & "~KoJD", vbExclamation, "Outdated"
If webRunWeb <> 0 Then Shell "explorer " & webRunWeb
End

End Sub

Public Sub CheckPWD()

Me.Hide
Form1.Hide
Form3.Show

End Sub

Public Sub ResumeIfOK()

If err_Name = False And err_Run = False And err_Ver = False And err_Pass = False Then
lblInit.Caption = "Initializing Succeeded!" & vbCrLf & "Starting..."
ProgressBar1.Value = 100
tmLatency.Enabled = True
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub tmLatency_Timer()
Form2.Hide
Form1.Show
tmLatency.Enabled = False
End Sub

