VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmMAPI 
   Caption         =   "Form1"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Previous"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtMsgNoteText 
      Height          =   4575
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmMAPI.frx":0000
      Top             =   2280
      Width           =   6855
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   6720
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   7920
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin VB.Label lblMsgSubject 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Label lblMsgOrigDisplayName 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   1440
      Width           =   4095
   End
   Begin VB.Label lblMsgDateReceived 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label lblMsgCount 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
End
Attribute VB_Name = "frmMAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objOutlook As Outlook.Application
Dim objMapiName As Outlook.NameSpace
Dim IntroComplete As Boolean
Dim LoadRequest(2)
Dim MailAgent As IAgentCtlCharacterEx
Dim Request As IAgentCtlRequest
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If MAPIMessages1.MsgIndex < MAPIMessages1.MsgCount - 1 Then
        MAPIMessages1.MsgIndex = MAPIMessages1.MsgIndex + 1
        DisplayMessage
    Else
        Beep
    End If

End Sub

Private Sub cmdPrevious_Click()

    If MAPIMessages1.MsgIndex > 0 Then
        MAPIMessages1.MsgIndex = MAPIMessages1.MsgIndex - 1
        DisplayMessage
    Else
        Beep
    End If

End Sub

Private Sub cmdSend_Click()

    With MAPIMessages1
        .MsgIndex = -1
        .RecipDisplayName = txtSendTo.Text
        .MsgSubject = txtSubject.Text
        .MsgNoteText = txtMessage.Text
        .SessionID = MAPISession1.SessionID
        .Send
    End With
    MsgBox "Message sent!", , "Send Message"

End Sub

Private Sub Form_Load()
    MAPISession1.SignOn
    MAPIMessages1.SessionID = MAPISession1.SessionID
    FetchNewMail
    DisplayMessage
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MAPISession1.SignOff
End Sub

Public Sub FetchNewMail()
    MAPIMessages1.FetchUnreadOnly = False
    MAPIMessages1.Fetch
End Sub

Public Sub DisplayMessage()
    lblMsgCount.Caption = "Message " & LTrim(Str(MAPIMessages1.MsgIndex + 1)) & " of " & LTrim(Str(MAPIMessages1.MsgCount))
    lblMsgDateReceived.Caption = MAPIMessages1.MsgDateReceived
    txtMsgNoteText.Text = MAPIMessages1.MsgNoteText
    lblMsgOrigDisplayName.Caption = MAPIMessages1.MsgOrigDisplayName
    lblMsgSubject.Caption = MAPIMessages1.MsgSubject
End Sub


