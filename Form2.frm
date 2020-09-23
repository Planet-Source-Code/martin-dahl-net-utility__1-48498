VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings..."
   ClientHeight    =   2625
   ClientLeft      =   8580
   ClientTop       =   7410
   ClientWidth     =   4350
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4350
   Begin VB.Frame Frame1 
      Caption         =   "Net Utility settings"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Scan services max port"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Timeout (default 50 ms)"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Winsock clones (default 20)"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim cl As Integer, ti As Integer, ra As Long
cl = Val(Text1.Text)
ti = Val(Text2.Text)
ra = Val(Text3.Text)

If cl < 1 Or ti < 1 Or ra < 1 Then
    MsgBox "Invalid input", vbCritical, "Error"
    Exit Sub
End If
Clones = cl
Timeout = ti
Range = ra
Form2.Hide
End Sub

Private Sub Command2_Click()
Form2.Hide
End Sub


Private Sub Form_Load()
Text1.Text = Clones
Text2.Text = Timeout
Text3.Text = Range
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii < 47 Or KeyAscii > 57 Then
    If KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii < 47 Or KeyAscii > 57 Then
    If KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub


