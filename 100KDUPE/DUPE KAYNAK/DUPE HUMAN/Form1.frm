VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "KOJD 100k DUPE 1745 !!!"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   1935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "DUPE"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1695
      Begin VB.TextBox txtInt 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "100"
         Top             =   960
         Width           =   1215
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   1320
         Top             =   840
      End
      Begin VB.CheckBox Check1 
         Caption         =   "START DUPE"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "GO"
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "knight online client"
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()

If Check1.Value = 1 Then
Timer1.Interval = CInt(txtInt.Text)
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If

End Sub

Private Sub Command1_Click()
LoadOffsets
If AttachKO = False Then
End
Exit Sub
End If
Me.Show
KO_ADR_CHR = ReadLong(KO_PTR_CHR)
KO_ADR_DLG = ReadLong(KO_PTR_DLG)
End Sub

Private Sub Timer1_Timer()

Dim pStr As String
Dim pBytes() As Byte

pStr = "6407b1010000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes

pStr = "55001031383030355F42696C626F722E6C7561"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes

pStr = "55001031383030355F42696C626F722E6C7561"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes

pStr = "55001031383030355F42696C626F722E6C7561"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes

pStr = "55001031383030355F42696C626F722E6C7561"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes

pStr = "6407EF260000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes

pStr = "55000F31333030395F4B756765722E6C7561"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes

pStr = "55000F31333030395F4B756765722E6C7561"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes

pStr = "55000F31333030395F4B756765722E6C7561"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes

pStr = "55000F31333030395F4B756765722E6C7561"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes

pStr = "6407b5010000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes

pStr = "640704270000"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes

pStr = "55001231313531305F466F726B7761696E2E6C7561"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes

pStr = "55001231313531305F466F726B7761696E2E6C7561"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes

End Sub
