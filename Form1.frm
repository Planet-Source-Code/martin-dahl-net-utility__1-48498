VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "-Net Utility- Version 1.02"
   ClientHeight    =   5610
   ClientLeft      =   7155
   ClientTop       =   5670
   ClientWidth     =   5865
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   5865
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5040
      Top             =   5760
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4200
      Top             =   5760
   End
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
      Left            =   6240
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   9551
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "IP information"
      TabPicture(0)   =   "Form1.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command8"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "OpenPorts"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "FadeAbout"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Startup"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command9"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command10"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Command11"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Scan network"
      TabPicture(1)   =   "Form1.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Command2"
      Tab(1).Control(2)=   "Command3"
      Tab(1).Control(3)=   "Command4"
      Tab(1).Control(4)=   "Command5"
      Tab(1).Control(5)=   "SSTab2"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Ping"
      TabPicture(2)   =   "Form1.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label9"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label10"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Text2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Command6"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "List1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Command12"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      Begin VB.CommandButton Command12 
         Caption         =   "Clear"
         Height          =   375
         Left            =   -74760
         TabIndex        =   47
         Top             =   1800
         Width           =   735
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Localhost"
         Height          =   255
         Left            =   3360
         TabIndex        =   46
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Send message to host"
         Height          =   255
         Left            =   960
         TabIndex        =   45
         Top             =   2520
         Width           =   2175
      End
      Begin VB.CommandButton Command9 
         Caption         =   "List"
         Height          =   375
         Left            =   4680
         TabIndex        =   44
         ToolTipText     =   "Show a list of the ports different services use"
         Top             =   2040
         Width           =   615
      End
      Begin VB.Timer Startup 
         Interval        =   10
         Left            =   5040
         Top             =   720
      End
      Begin VB.Timer FadeAbout 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   5040
         Top             =   1320
      End
      Begin VB.ListBox OpenPorts 
         Appearance      =   0  'Flat
         Height          =   1980
         Left            =   360
         TabIndex        =   43
         Top             =   3000
         Width           =   4815
         Visible         =   0   'False
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Scan services"
         Height          =   375
         Left            =   3120
         TabIndex        =   42
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Copy to Scan"
         Height          =   375
         Left            =   1560
         TabIndex        =   41
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ListBox List1 
         Height          =   3765
         ItemData        =   "Form1.frx":035E
         Left            =   -73800
         List            =   "Form1.frx":0360
         TabIndex        =   36
         Top             =   1320
         Width           =   4095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Start"
         Height          =   325
         Left            =   -70680
         TabIndex        =   35
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   -73800
         TabIndex        =   32
         Top             =   720
         Width           =   2895
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   29
         Top             =   2880
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4260
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Subnet listings"
         TabPicture(0)   =   "Form1.frx":0362
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "TreeView1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Hits"
         TabPicture(1)   =   "Form1.frx":037E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TreeView2"
         Tab(1).ControlCount=   1
         Begin MSComctlLib.TreeView TreeView2 
            Height          =   1815
            Left            =   -74880
            TabIndex        =   31
            Top             =   480
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   3201
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Appearance      =   1
         End
         Begin MSComctlLib.TreeView TreeView1 
            CausesValidation=   0   'False
            Height          =   1815
            Left            =   120
            TabIndex        =   30
            Top             =   480
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   3201
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Appearance      =   1
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Settings"
         Height          =   375
         Left            =   -70800
         TabIndex        =   28
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Clear"
         Height          =   375
         Left            =   -72120
         TabIndex        =   25
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Stop"
         Height          =   375
         Left            =   -73440
         TabIndex        =   24
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Start"
         Height          =   375
         Left            =   -74760
         TabIndex        =   23
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Caption         =   "Range"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   5295
         Begin VB.TextBox IP2 
            Height          =   285
            Index           =   3
            Left            =   4680
            MaxLength       =   3
            TabIndex        =   16
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox IP2 
            Height          =   285
            Index           =   2
            Left            =   4200
            MaxLength       =   3
            TabIndex        =   15
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox IP2 
            Height          =   285
            Index           =   1
            Left            =   3720
            MaxLength       =   3
            TabIndex        =   14
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox IP1 
            Height          =   285
            Index           =   3
            Left            =   2040
            MaxLength       =   3
            TabIndex        =   12
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox IP1 
            Height          =   285
            Index           =   2
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   11
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox IP1 
            Height          =   285
            Index           =   1
            Left            =   1080
            MaxLength       =   3
            TabIndex        =   10
            Top             =   480
            Width           =   375
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Single port only"
            Height          =   255
            Left            =   600
            TabIndex        =   26
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Show domain names (slow)"
            Height          =   255
            Left            =   2760
            TabIndex        =   27
            Top             =   1440
            Width           =   2415
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   3240
            MaxLength       =   5
            TabIndex        =   22
            Text            =   "10"
            Top             =   960
            Width           =   1935
            Visible         =   0   'False
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   600
            MaxLength       =   5
            TabIndex        =   21
            Text            =   "1"
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox IP2 
            Height          =   285
            Index           =   0
            Left            =   3240
            MaxLength       =   3
            TabIndex        =   13
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox IP1 
            Height          =   285
            Index           =   0
            Left            =   600
            MaxLength       =   3
            TabIndex        =   9
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label6 
            Caption         =   "To"
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   20
            Top             =   960
            Width           =   375
            Visible         =   0   'False
         End
         Begin VB.Label Label8 
            Caption         =   "Port"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label7 
            Caption         =   "IP"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label6 
            Caption         =   "To"
            Height          =   255
            Index           =   0
            Left            =   2760
            TabIndex        =   17
            Top             =   480
            Width           =   375
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         MaxLength       =   32
         TabIndex        =   1
         Top             =   600
         Width           =   3015
      End
      Begin VB.Frame Frame2 
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   240
         TabIndex        =   37
         Top             =   2760
         Width           =   5055
         Begin VB.Label Label12 
            Caption         =   "By clicking on either the information field or the name filend you can add the informating to the clipboard."
            Height          =   735
            Left            =   360
            TabIndex        =   39
            Top             =   1320
            Width           =   4095
         End
         Begin VB.Label Label11 
            Caption         =   "This program was created by Martin Dahl in 2003. You can reach me on calypso@hotbrev.com"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   360
            TabIndex        =   38
            Top             =   480
            Width           =   3975
         End
      End
      Begin VB.Label Label10 
         Caption         =   "Response"
         Height          =   255
         Left            =   -74760
         TabIndex        =   34
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Host"
         Height          =   255
         Left            =   -74760
         TabIndex        =   33
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "Canonical Name"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Information"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Host"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   40
      Top             =   5640
      Width           =   5175
   End
   Begin VB.Menu menu 
      Caption         =   "Main Menu"
      Visible         =   0   'False
      Begin VB.Menu Main_menu 
         Caption         =   "Export entire list to clipboard"
         Index           =   0
      End
      Begin VB.Menu Main_menu 
         Caption         =   "Export list item to clipboard"
         Index           =   1
      End
   End
   Begin VB.Menu menu2 
      Caption         =   "OpenPortMenu"
      Visible         =   0   'False
      Begin VB.Menu Main_menu2 
         Caption         =   "Hide list"
         Index           =   0
      End
      Begin VB.Menu Main_menu2 
         Caption         =   "Copy to clipboard"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Public Function getRandomKey(length As Integer) As String
Dim rstring As String
Dim i As Integer, RandomNumber As Integer



For i = 1 To length
    RandomNumber = Int((122 - 97 + 1) * Rnd + 97)
    rstring = rstring & Chr(RandomNumber)
Next

getRandomKey = rstring

End Function


Function IPCheck(inp As String) As Boolean
Dim IPpart(0 To 3) As Integer
Dim i As Integer, ascPart As Integer, IPartCnt As Integer
Dim pos As Integer
On Error Resume Next
If Len(inp) > 6 And Len(inp) < 16 Then
    For i = 1 To Len(inp)
        ascPart = Asc(Mid(inp, i, 1))
        If (ascPart > 57 Or ascPart < 48) Then
            If ascPart <> 46 Then Exit Function
        End If
    Next
    
    i = 1
    While i > 0
        i = InStr(1, inp, ".", vbBinaryCompare)
        If i = 1 Or IPartCnt > 3 Then Exit Function
        If IPartCnt = 3 Then
            IPpart(IPartCnt) = Val(inp)
        Else
            IPpart(IPartCnt) = Val(Mid(inp, 1, i - 1))
        End If
        IPartCnt = IPartCnt + 1
        inp = Mid(inp, i + 1, Len(inp))
    Wend
    For i = 0 To 3
        If IPpart(i) > 255 Then Exit Function
    Next
    
    
    IPCheck = True
Else
    IPCheck = False
End If



End Function



Private Sub Check1_Click()
'If Check1.Value = 1 Then
'    Text4.Enabled = False
'    Text5.Enabled = False
'Else
'    Text4.Enabled = True
'    Text5.Enabled = True
'End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Label6(1).Visible = False
    Text5.Visible = False
Else
    Label6(1).Visible = True
    Text5.Visible = True
End If


End Sub

Private Sub Command1_Click()
Dim Hostent As Hostent, host As String
Dim PointerToPointer As Long, ListAddr As Long
Dim IPLong As Inet_address
Dim IP As String
Dim szString As String

host = Trim$(Text1.Text)
szString = String(64, &H0)
host = host + Right$(szString, 64 - Len(host))


If Len(Text1.Text) < 4 Then Exit Sub

If IPCheck(Text1.Text) Then
    Label2.Caption = Text1.Text
    Label5.Caption = vbgetHostName(Text1.Text)
Else
    If gethostbyname(host) = SOCKET_ERROR Then
        MsgBox "Can not resolve host", vbOKOnly, "ws error"
    Else
        PointerToPointer = gethostbyname(host) ' Get the pointer to the address of the winsock hostent structure
        CopyMemory Hostent, ByVal PointerToPointer, Len(Hostent)  ' Copy Winsock structure to the VisualBasic structure
        CopyMemory ListAddr, ByVal Hostent.h_addr_list, 4 ' Copy Winsock structure To the VisualBasic structure
        CopyMemory IPLong, ByVal ListAddr, 4 ' Get the first list entry from the Address List
        IP = Trim$(CStr(Asc(IPLong.Byte4)) + "." + CStr(Asc(IPLong.Byte3)) _
        + "." + CStr(Asc(IPLong.Byte2)) + "." + CStr(Asc(IPLong.Byte1)))
        Label5.Caption = Text1.Text
        Label2.Caption = IP
    End If

End If
Text2.Text = Label2.Caption
End Sub

Private Sub Command10_Click()
Dim cmdstr As String
Dim message As String
message = InputBox("Enter message to send", "message?")
If message = "" Or Label2.Caption = Empty Then
    MsgBox "invalid input", vbCritical, " error"
    Exit Sub
End If
cmdstr = "net send " & Label2.Caption & " " & message

Call Shell(cmdstr, vbHide)
End Sub

Private Sub Command11_Click()
Text1.Text = Winsock(0).LocalIP
Command1_Click
End Sub


Private Sub Command12_Click()
List1.Clear
End Sub

Private Sub Command2_Click()
Dim TheNode As Node
Dim ProbingIP(0 To 3) As Integer
Dim EndIP(0 To 3) As Integer, ws As Integer
Dim key As String, RootEnum As String, ipn As String
Dim WsHistory() As String
Dim hits As Integer
Dim host As String
On Error Resume Next

ReDim WsHistory(0 To Clones - 1)

Timer1.Enabled = True
If Running Then Exit Sub
Running = True
RootEnum = "Root 1"
key = "c 1"
ws = 0

If Winsock.Count <> Clones Then
    For i = 0 To Clones
        If i > Winsock.UBound Then Load Winsock(Winsock.UBound + 1)
        Winsock(i).Close
        DoEvents
    Next
End If


For i = 0 To 3
    If IP1(i) > 255 Or IP2(i) > 255 Then
        MsgBox "Invalid input", vbCritical, "error"
        Exit Sub
    End If
    ProbingIP(i) = IP1(i)
    EndIP(i) = IP2(i)
Next

TreeView1.Nodes.Clear
TreeView2.Nodes.Clear

Set TheNode = TreeView1.Nodes.Add(, , RootEnum, IP1(0) & "." & IP1(1) & "." & IP1(2) & "." & IP1(3))
Set TheNode = TreeView2.Nodes.Add(, , "Port Root", "Port: " & Text4.Text)

Do While ProbingIP(2) <= EndIP(2)
DoEvents

Do While ProbingIP(3) <= EndIP(3) And Running

     ipn = ProbingIP(0) & "." & ProbingIP(1) & "." & ProbingIP(2) & "." & ProbingIP(3)
     Label13.Caption = "Status: Scaining " & ipn
     
     If Check1.Value = 0 Then
        Set TheNode = TreeView1.Nodes.Add(RootEnum, tvwChild, key, ipn)
    Else
        Label13.Caption = "Status: Scaning ip " & ipn
        host = vbgetHostName(ipn)
        If Len(host) > 0 Then
            Set TheNode = TreeView1.Nodes.Add(RootEnum, tvwChild, key, ipn & "<==>" & host)
        Else
            Set TheNode = TreeView1.Nodes.Add(RootEnum, tvwChild, key, ipn)
        End If
        DoEvents
        DoEvents
        DoEvents
    End If
     
     If ws < Clones Then
        WsHistory(ws) = ipn
        Winsock(ws).Connect ipn, Text4.Text
        ws = ws + 1
     Else
        Sleep (Timeout)
        ws = 0
        For i = 0 To Winsock.UBound
            If Winsock(i).State = 7 Then
                Set TheNode = TreeView2.Nodes.Add("Port Root", tvwChild, "p " & hits, WsHistory(i))
                hits = hits + 1
            End If
            Winsock(i).Close
        Next
    End If
    
    
    key = "c" & Str(Val(Mid(key, 2, Len(key)) + 1))
    'If ProbingIP(3) < 255 Then
        ProbingIP(3) = ProbingIP(3) + 1
    'End If
    DoEvents
    
Loop

For i = 0 To Winsock.UBound
    If Winsock(i).State = 7 Then
        Set TheNode = TreeView2.Nodes.Add("Port Root", tvwChild, "p " & hits, WsHistory(i))
        hits = hits + 1
    End If
    Winsock(i).Close
Next

If ProbingIP(2) < EndIP(2) And Running Then
    ProbingIP(3) = 0
    RootEnum = "Root" & Str(Val(Mid(RootEnum, 6, Len(key)) + 1))
    ProbingIP(2) = ProbingIP(2) + 1
    Set TheNode = TreeView1.Nodes.Add(, , RootEnum, ProbingIP(0) & "." & ProbingIP(1) & "." & ProbingIP(2) & "." & ProbingIP(3))
    
Else
    Exit Do
End If


Loop

Label13.Caption = "Status:"
Timer2.Enabled = True
'TheNode.EnsureVisible

End Sub

Private Sub Command3_Click()
Running = False
Timer1.Enabled = False
Timer2.Enabled = True
End Sub

Private Sub Command4_Click()
TreeView1.Nodes.Clear
TreeView2.Nodes.Clear

End Sub


Private Sub Command5_Click()
Form2.Show
End Sub

Private Sub Command6_Click()
Dim result As Long
Dim ECHO As ICMP_ECHO_REPLY

Call Ping(Text2.Text, ECHO)
result = ECHO.RoundTripTime
If ECHO.status = 0 Then
    List1.AddItem Text2.Text & vbTab & result & " ms"
Else
    MsgBox "destination unreachable", vbCritical, "Winsock error"
End If
End Sub

Private Sub Command7_Click()
Dim i As Integer, tmp As String, s As Integer
Dim IP(0 To 3) As Integer
i = 0
s = 1
If Label2.Caption = Empty Then Exit Sub
tmp = Label2.Caption & "."


Do While i < 3
    If InStr(s, tmp, ".", vbBinaryCompare) > 0 Then
        t = InStr(s, tmp, ".", vbBinaryCompare)
        IP(i) = Mid(tmp, s, t - s)
        s = t + 1
        i = i + 1
    Else
        Exit Do
    End If
'DoEvents
Loop

For i = 0 To 3
    IP1(i) = IP(i)
    IP2(i) = IP(i)
Next
IP2(3) = 255
End Sub

Private Sub Command8_Click()
Dim i As Integer, b As Integer, pos As Integer
Dim IP As String, ws As Integer
On Error Resume Next

If Text1.Text = Empty Then Exit Sub

If Command8.Caption = "Scan services" Then
    Command8.Caption = "Cancel Scan"
    OpenPorts.Visible = True
    FadeAbout.Interval = 1
    FadeAbout.Tag = 0
    FadeAbout.Enabled = True
Else
    Command8.Caption = "Scan services"
    Running = False
    Exit Sub
End If

Running = True

Timer1.Enabled = True

ws = 0
OpenPorts.Clear


IP = Text1.Text
If Clones / 2 > Range Then Clones = Int(Range / 2)

If Winsock.Count <> Clones Then
    For i = 0 To Clones
        If i > Winsock.UBound Then Load Winsock(Winsock.UBound + 1)
        Winsock(i).Close
        DoEvents
    Next
End If

For i = 1 To Range
DoEvents
If Not Running Then Exit For
    If i < Clones + pos Then
        Winsock(ws).Close
        DoEvents
        Winsock(ws).Connect IP, i
        ws = ws + 1
    Else
        pos = pos + Clones
        ws = 0
        Sleep (Timeout)
        For b = i To i + Clones
            If Winsock(ws).State = 7 Then
            OpenPorts.AddItem Winsock(ws).RemotePort
            End If
            ws = ws + 1
        Next
        Label13.Caption = "Status: Scaning port: " & i
        ws = 0
    End If


Next
For i = Winsock.LBound To Winsock.Count - 1
    If Winsock(i).State = 7 Then
        OpenPorts.AddItem Winsock(i).RemotePort
    End If
    Winsock(i).Close
Next

Label13.Caption = "Status:"
Timer2.Enabled = True
Running = False
Command8.Caption = "Scan services"

End Sub




Private Sub Command9_Click()
Form3.Show
End Sub

Private Sub FadeAbout_Timer()
FadeAbout.Tag = FadeAbout.Tag + 1

FadeAbout.Interval = FadeAbout.Interval + 8
If Frame2.Visible Then
    Frame2.Visible = False
Else
    Frame2.Visible = True
End If
If FadeAbout.Tag > 20 Then FadeAbout.Enabled = False

End Sub

Private Sub Form_Load()
Dim tmp As String
Dim IP(0 To 3) As Integer
Dim i As Integer, s As Integer, t As Integer
Transparent Me.hWnd, 0
i = 0
s = 1
Clones = 20
Timeout = 50
Range = 5000

'Randomize
Text1.Text = Winsock(0).LocalIP
tmp = Winsock(0).LocalIP & "."

Do While i < 3
    If InStr(s, tmp, ".", vbBinaryCompare) > 0 Then
        t = InStr(s, tmp, ".", vbBinaryCompare)
        IP(i) = Mid(tmp, s, t - s)
        s = t + 1
        i = i + 1
    Else
        Exit Do
    End If
'DoEvents
Loop

For i = 0 To 3
    IP1(i) = IP(i)
    IP2(i) = IP(i)
Next

IP2(3) = 255



End Sub

Private Sub Form_Unload(Cancel As Integer)
For i = 0 To Winsock().UBound
    Winsock(i).Close
Next
Unload Form2
Unload Form3
End
End Sub

Private Sub IP1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii < 47 Or KeyAscii > 57 Then
    If KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub


Private Sub IP2_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii < 47 Or KeyAscii > 57 Then
    If KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub


Private Sub Label2_Click()
Clipboard.Clear
Clipboard.SetText Label2.Caption
End Sub

Private Sub Label5_Click()
Clipboard.Clear
Clipboard.SetText Label5.Caption
End Sub


Private Sub Main_menu_Click(Index As Integer)
Dim i As Integer
Dim tmp As String
If TreeView1.Nodes.Count = 0 Then Exit Sub
Clipboard.Clear
Select Case Index
Case 0:
    For i = 1 To TreeView1.Nodes.Count
        tmp = tmp & TreeView1.Nodes.Item(i).Text & vbNewLine
    Next
    Clipboard.SetText tmp, vbCFText
Case 1:
    Clipboard.SetText TreeView1.SelectedItem.Text, vbCFText
Case Default:
End Select

End Sub

Private Sub Main_menu2_Click(Index As Integer)
Dim i As Integer, tmp As String
Select Case Index
Case 0:
    OpenPorts.Visible = False
    Frame2.Visible = True
Case 1
    For i = 0 To OpenPorts.ListCount
        tmp = tmp & OpenPorts.List(i) & vbNewLine
    Next
Case Default:
End Select
Clipboard.Clear
Clipboard.SetText tmp
End Sub

Private Sub OpenPorts_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu menu2
End If
End Sub


Private Sub Startup_Timer()
Static fade
Dim tmp As Integer

fade = fade + 2
tmp = Int(fade)
Transparent Me.hWnd, tmp
If fade >= 255 Then
    Startup.Enabled = False
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1_Click
End Sub


Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii < 47 Or KeyAscii > 57 Then
    If KeyAscii <> 8 Then KeyAscii = 0
End If


End Sub


Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii < 47 Or KeyAscii > 57 Then
    If KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub








Private Sub Timer1_Timer()

If Form1.Height < 6450 And Timer2.Enabled = False Then
    If Form1.WindowState = vbNormal Then Form1.Height = Form1.Height + 20
Else
    Timer1.Enabled = False
End If

End Sub

Private Sub Timer2_Timer()
If Form1.Height > 6090 Then
    If Form1.WindowState = vbNormal Then Form1.Height = Form1.Height - 20
Else
    Timer2.Enabled = False
End If

End Sub


Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu menu, vbPopupMenuLeftAlign
End If
End Sub


