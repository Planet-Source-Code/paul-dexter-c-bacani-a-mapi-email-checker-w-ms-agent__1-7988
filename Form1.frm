VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuration..."
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5220
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   6000
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   5280
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   120
      Pattern         =   "*.acs"
      TabIndex        =   6
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1095
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   350
      Left            =   4800
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   300
   End
   Begin VB.TextBox txtInterval 
      Height          =   330
      Left            =   4440
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5280
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5280
      Top             =   120
   End
   Begin VB.Label lblCap 
      BackStyle       =   0  'Transparent
      Caption         =   "Interval (minutes) to check mail"
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   3
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblCap 
      BackStyle       =   0  'Transparent
      Caption         =   "MS Agents"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin AgentObjectsCtl.Agent Agent 
      Left            =   5280
      Top             =   1320
      _cx             =   847
      _cy             =   847
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub FillAgents(AgentDir As String)
    On Error Resume Next
    File1.Path = AgentDir
    LoopCnt = 0
    Do While File1.ListCount = 0
        LoopCnt = LoopCnt + 1
        AgentDir = InputBox$("Please enter directory where MS Agents characters where installed.", "Agent Directory", "")
        File1.Path = AgentDir
        If LoopCnt = 3 Then
            MsgBox "No MS Agent installed on your system.", , "Error"
            End
            Exit Sub
        End If
    Loop
    TmpAgent = GetIni("MS Agent", "Character", App.Path & "\EmailChk.ini")
    If TmpAgent = "" Then
        x = WriteIni("MS Agent", "Character", UCase$(Left$(File1.List(0), InStr(File1.List(0), ".") - 1)), App.Path & "\EmailChk.ini")
    Else
        For i = 0 To File1.ListCount - 1
            If UCase$(Left$(File1.List(i), Len(TmpAgent))) = TmpAgent Then
                File1.ListIndex = i
                Exit For
            End If
        Next i
    End If
    File1.Tag = File1.ListIndex
End Sub

Private Sub Agent_Command(ByVal UserInput As Object)
    On Error Resume Next
    Select Case UserInput.Name
        Case "Configuration"
            Form1.Show
        Case "Check"
            Form1.Tag = 0
            MessageCount = 0
            Callme
        Case "Read"
            Form2.Show
        Case "Exit"
            MailAgent.Stop
            MailAgent.Show
            MailAgent.Play "Wave"
            MailAgent.Speak "Thank you for using Outlook email checker"
            MailAgent.Play "Respose"
            DoEvents: DoEvents
            Set LoadRequest(1) = MailAgent.Hide
        Case "Time"
            mHour = Format$(Time, "h")
            mMinute = Format$(Time, "m")
            m = Format$(Time, "am/pm")
            MailAgent.Stop
            MailAgent.Show
            MailAgent.Speak "It is now " & mHour & ":" & mMinute & "\pau=10\ " & m
            DoEvents
            MailAgent.Hide
     End Select
End Sub

Private Sub Agent_RequestComplete(ByVal Request As Object)
    On Error Resume Next
    Select Case Request
        Case LoadRequest(0)
            Set MailAgent = Agent.Characters(UCase$(Left$(File1.List(File1.ListIndex), InStr(File1.List(File1.ListIndex), ".") - 1)))
        Case LoadRequest(1)
            Agent.Characters.Unload UCase$(Left$(File1.List(File1.ListIndex), InStr(File1.List(File1.ListIndex), ".") - 1))
            Unload Me
    End Select
End Sub

Private Sub cmdCancel_Click()
    File1.ListIndex = File1.Tag
    txtInterval.Text = txtInterval.Tag
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    x = WriteIni("Email Checker", "TimerInterval", txtInterval, App.Path & "\EmailChk.ini")
    txtInterval.Tag = txtInterval.Text
    x = WriteIni("MS Agent", "Character", UCase$(Left$(File1.List(File1.ListIndex), InStr(File1.List(File1.ListIndex), ".") - 1)), App.Path & "\EmailChk.ini")
    If File1.Tag <> File1.ListIndex Then
        Agent.Characters.Unload UCase$(Left$(File1.List(File1.Tag), InStr(File1.List(File1.Tag), ".") - 1))
        Set LoadRequest(0) = Agent.Characters.Load(UCase$(Left$(File1.List(File1.ListIndex), InStr(File1.List(File1.ListIndex), ".") - 1)), File1.List(File1.ListIndex))
        File1.Tag = File1.ListIndex
        Form1.Tag = 0
        Timer2.Enabled = True
    End If
    Form1.Hide
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorLoad
    Dim wDir As String
    MAPISession1.SignOn
    MAPIMessages1.SessionID = MAPISession1.SessionID
    Form1.Tag = 0
    'loads the MS Agent
    wDir = WinDir()
    If Right$(wDir, 1) = "\" Then wDir = wDir & "MSAgent\Chars" Else wDir = wDir & "\MSAgent\Chars"
    FillAgents wDir
    
    txtInterval = GetIni("Email Checker", "TimerInterval", App.Path & "\EmailChk.ini")
    If txtInterval = "" Then
        txtInterval = 15
        x = WriteIni("Email Checker", "TimerInterval", txtInterval, App.Path & "\EmailChk.ini")
    End If
    txtInterval.Tag = txtInterval.Text
    VScroll1.Value = txtInterval
    
    Set LoadRequest(0) = Agent.Characters.Load(UCase$(Left$(File1.List(File1.ListIndex), InStr(File1.List(File1.ListIndex), ".") - 1)), File1.List(File1.ListIndex))
    'Timer2 is to give time to the MS Agent to load
    'There must be some other way to do this. I'm sure you can do it.
    Timer2.Enabled = True
    'Saves the last number of mails found.
    MessageCount = 0
    'is use on the Timer1 - Counter multiply by the number of intervals
    'is the amount of time to check for mail
    Counter = 0
    Exit Sub
ErrorLoad:
    MsgBox Err.Description
    MAPISession1.SignOff
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MAPISession1.SignOff
    End
End Sub
Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Counter = Counter + 1
    mHour = Format$(Time, "h")
    mMinute = Format$(Time, "m")
    m = Format$(Time, "am/pm")
    If mMinute = "00" Or mMinute = "30" Then
        MailAgent.Stop
        MailAgent.Show
        MailAgent.Play "It is now " & mHour & "\pau=25\" & mMinute & "\pau=15\" & m
        DoEvents
        MailAgent.Hide
    End If
    If Counter = txtInterval.Tag Then
        Callme
        Counter = 0
    End If
    Timer1.Enabled = True
    DoEvents
End Sub
Sub Callme()
    Dim intCountUnRead As Integer, intCountTotal As Integer
    On Error Resume Next
    'Check unread mail on your inbox Folder
    'This code I adapted from Tim Ford
    MAPIMessages1.FetchUnreadOnly = True
    MAPIMessages1.Fetch
    intCountUnRead = MAPIMessages1.MsgCount
    MAPIMessages1.FetchUnreadOnly = False
    MAPIMessages1.Fetch
    intCountTotal = MAPIMessages1.MsgCount
    'Show MS Agent if new mail is found
    If (intCountUnRead > 0 And MessageCount <> intCountUnRead) Or Form1.Tag = 0 Then
        MailAgent.Stop
        MailAgent.Show
        MailAgent.Play "Uncertain"
        'No mail
        If MessageCount = 0 Then
            If intCountUnRead = 0 Then
                MailAgent.Play "Sad"
                MailAgent.Speak "Sorry, You have no\pau=100\ new \pau=50\e-mail message"
            Else
                If intCountUnRead = 1 Then msg = "message" Else msg = "messages"
                MailAgent.Play "Congratulate"
                MailAgent.Speak "You have " & intCountUnRead & "\pau=100\unread \pau=50\e-mail " & msg
            End If
        Else
            'intCountRead is the number of unread mails
            If intCountUnRead = 1 Then msg = "message" Else msg = "messages"
            MailAgent.Play "Congratulate"
            MailAgent.Speak "You now have " & intCountUnRead & "\pau=100\unread \pau=50\e-mail " & msg
        End If
        MailAgent.Play "RestPose"
        MailAgent.Hide
        DoEvents
        MessageCount = intCountUnRead
    End If
    
    intCountUnRead = 0
    
    Form1.Tag = 1
    
End Sub

Private Sub Timer2_Timer()
    On Error Resume Next
    Timer2.Enabled = False
    MailAgent.Commands.Add "Time", "Time", "Time", True, True
    MailAgent.Commands.Add "Configuration", "Configuration", "Configuration", True, True
    MailAgent.Commands.Add "Check", "Check Email", "Check Email", True, True
    MailAgent.Commands.Add "Read", "Read Mail", "Read Mail", True, True
    MailAgent.Commands.Add "Exit", "Exit", "Exit", True, True
    MailAgent.Commands.Caption = "MailAgent"
    Callme
    Timer1.Enabled = True
End Sub

Private Sub txtInterval_GotFocus()
    txtInterval.SelLength = Len(txtInterval)
    txtInterval.SelStart = 0
End Sub

Private Sub VScroll1_Change()
    If VScroll1.Value = 4 Then
        VScroll1.Value = txtInterval
        Exit Sub
    End If
    txtInterval = VScroll1.Value
End Sub
