Attribute VB_Name = "Module1"
Global IntroComplete As Boolean
Global LoadRequest(2)
Global MailAgent As IAgentCtlCharacterEx
Global Request As IAgentCtlRequest

Global MessageCount As Integer, LogUser As String
Global Counter As Integer
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Function GetUser() As String


    Dim lpUserID As String
    Dim nBuffer As Long
    Dim Ret As Long
    lpUserID = String(25, 0)
    nBuffer = 25
    Ret = GetUserName(lpUserID, nBuffer)
    If Ret Then
        GetUser$ = ClipNull(lpUserID$)
    End If


End Function


Function ClipNull(InString As String) As String


    Dim intpos As Integer
    If Len(InString) Then
        intpos = InStr(InString, vbNullChar)
        If intpos > 0 Then
            ClipNull = Left(InString, intpos - 1)
        Else
            ClipNull = InString
        End If


    End If


End Function


Function WinDir() As String
    Dim TmpDir As String * 255
    Dim i As Long
    i = GetWindowsDirectory(TmpDir, Len(TmpDir))
    WinDir = Left(TmpDir, i)
End Function

Function GetIni(strSectionHeader As String, strVariableName As String, strFileName As String) As String
    Dim strReturn As String
    ' Blank the return string
    strReturn = String(255, Chr(0))
    'Get requested information, trimming the
    '     returned
    ' string
    GetIni = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFileName))
End Function


Function WriteIni(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer
    WriteIni = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)
End Function
