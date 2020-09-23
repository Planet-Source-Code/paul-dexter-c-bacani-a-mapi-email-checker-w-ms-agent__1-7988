VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "E-mails...."
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   7320
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >>"
      Height          =   495
      Left            =   7320
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<< Previous"
      Height          =   495
      Left            =   7320
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtMsgNoteText 
      Height          =   4695
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2280
      Width           =   8535
   End
   Begin VB.Label lblMsgSubject 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   6495
   End
   Begin VB.Label lblMsgOrigDisplayName 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   6495
   End
   Begin VB.Label lblMsgDateReceived 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label lblMsgCount 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const BalloonOn = 1
Private Sub cmdClose_Click()
    MailAgent.Stop
    MailAgent.Hide
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If Form1.MAPIMessages1.MsgIndex < Form1.MAPIMessages1.MsgCount - 1 Then
        Form1.MAPIMessages1.MsgIndex = Form1.MAPIMessages1.MsgIndex + 1
        DisplayMessage
    Else
        Beep
    End If

End Sub

Private Sub cmdPrevious_Click()

    If Form1.MAPIMessages1.MsgIndex > 0 Then
        Form1.MAPIMessages1.MsgIndex = Form1.MAPIMessages1.MsgIndex - 1
        DisplayMessage
    Else
        Beep
    End If

End Sub


Private Sub Form_Load()
    MailAgent.Show
    MailAgent.Balloon.Style = MailAgent.Balloon.Style And (Not BalloonOn)
    FetchNewMail
    DisplayMessage
End Sub

Public Sub FetchNewMail()
    Form1.MAPIMessages1.FetchUnreadOnly = False
    Form1.MAPIMessages1.Fetch
End Sub

Public Sub DisplayMessage()
    On Error Resume Next
    lblMsgCount.Caption = "Message " & LTrim(Str(Form1.MAPIMessages1.MsgIndex + 1)) & " of " & LTrim(Str(Form1.MAPIMessages1.MsgCount))
    lblMsgDateReceived.Caption = Format$(Form1.MAPIMessages1.MsgDateReceived, "dd mmmm yyyy,  hh:mm am/pm")
    txtMsgNoteText.Text = Form1.MAPIMessages1.MsgNoteText
    lblMsgOrigDisplayName.Caption = Form1.MAPIMessages1.MsgOrigDisplayName
    lblMsgSubject.Caption = Form1.MAPIMessages1.MsgSubject
    MailAgent.Stop
    MailAgent.Play "Read"
    MailAgent.Speak "This is a mail from " & Form1.MAPIMessages1.MsgOrigDisplayName
    MailAgent.Speak txtMsgNoteText.Text
    MailAgent.Play Choose(Int(8 * Rnd) + 1, "Idle1_1", "Idle1_2", "Idle1_3", "Idle1_4", "Idle2_1", "Idle2_2", "Idle3_1", "Idle3_2")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MailAgent.Balloon.Style = MailAgent.Balloon.Style Or BalloonOn
End Sub

