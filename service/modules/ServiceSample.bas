Attribute VB_Name = "service"
Option Explicit

Public Const INFINITE = -1&      '  Infinite timeout
Private Const WAIT_TIMEOUT = 258&
Public teller As Integer

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(1 To 128) As Byte      '  Maintenance string for PSS usage
End Type

Public Const VER_PLATFORM_WIN32_NT = 2&

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Public hStopEvent As Long, hStartEvent As Long, hStopPendingEvent
Public IsNT As Boolean, IsNTService As Boolean
Public ServiceName() As Byte, ServiceNamePtr As Long
Private Declare Function CmdAsUser Lib "cmddll.dll" (ByVal uUsercred As String, ByVal uUserPass As String, ByVal sCmdLIne As String, ByVal sStartDir As String, ByVal WindowMode As Long) As Long


Private Sub Main()
Dim go_file As String
Dim cmd_file As String
Dim run_cmd As String
Dim found As String

Dim inter_val As Integer
Dim foreground As String
Dim xx As Integer
Dim appa As String

Dim hnd As Long
Dim h(0 To 1) As Long
    
    
If LCase(Command) = "" Then frmServiceControl.Show
If LCase(Command) = "-bug" Then form1.Show
If LCase(Command) = "-service_start" Then



    ' Check OS type
    IsNT = CheckIsNT()
    ' Creating events
    hStopEvent = CreateEvent(0, 1, 0, vbNullString)
    hStopPendingEvent = CreateEvent(0, 1, 0, vbNullString)
    hStartEvent = CreateEvent(0, 1, 0, vbNullString)
    ServiceName = StrConv(frmServiceControl.SERVICE_NAME, vbFromUnicode)
    ServiceNamePtr = VarPtr(ServiceName(LBound(ServiceName)))
    If IsNT Then
        ' Trying to start service
        hnd = StartAsService
        h(0) = hnd
        h(1) = hStartEvent
        ' Waiting for one of two events: sucsessful service start (1) or
        ' terminaton of service thread (0)
        IsNTService = WaitForMultipleObjects(2&, h(0), 0&, INFINITE) = 1&
        If Not IsNTService Then
          CloseHandle hnd
            'MsgBox "This program must be started as service."
            ''MessageBox 0&, "This program must be started as a service.", App.Title, vbInformation Or vbOKOnly Or vbMsgBoxSetForeground
            frmServiceControl.Show
            Exit Sub
            
        End If
    Else
        MessageBox 0&, "This program is only for Windows NT/2000/XP.", app.Title, vbInformation Or vbOKOnly Or vbMsgBoxSetForeground
       ''     frmServiceControl.Show
            Exit Sub
    End If

    If IsNTService Then
inter_val = Val(GetFromINI("service", "interval", app.Path & "\service.ini"))

cmd_file = GetFromINI("cmd", "cmd_file", app.Path & "\service.ini")
foreground = GetFromINI("service", "fore_ground", app.Path & "\service.ini")
If foreground = "1" Then ToTaskTray frmServiceControl


        SetServiceState SERVICE_RUNNING
        app.LogEvent frmServiceControl.SERVICE_NAME + " started"
        Do



teller = teller - 1
If teller <= 0 Then

If foreground = "1" Then ReloadTray frmServiceControl




If Dir(cmd_file) <> "" Then


run_cmd = GetFromINI("command", "Shell", cmd_file)
If Dir(run_cmd) <> "" Then Shell run_cmd, vbNormalFocus


Kill (cmd_file)
End If
fill_boxes
Call FillProcessListNT(frmServiceControl.List1)
run_tasks
kill_tasks
teller = inter_val
End If





        Loop While WaitForSingleObject(hStopPendingEvent, 1000&) = WAIT_TIMEOUT

        SetServiceState SERVICE_STOPPED
        app.LogEvent frmServiceControl.SERVICE_NAME + " stopped"
        SetEvent hStopEvent
        ' Waiting for service thread termination
        WaitForSingleObject hnd, INFINITE
        CloseHandle hnd
        UnloadTray
    End If
    CloseHandle hStopEvent
    CloseHandle hStartEvent
    CloseHandle hStopPendingEvent
    
    End If
    
End Sub

' CheckIsNT() returns True, if the program runs
' under Windows NT or Windows 2000, and False
' otherwise.
Public Function CheckIsNT() As Boolean
    Dim OSVer As OSVERSIONINFO
    OSVer.dwOSVersionInfoSize = LenB(OSVer)
    GetVersionEx OSVer
    CheckIsNT = OSVer.dwPlatformId = VER_PLATFORM_WIN32_NT
End Function

Public Function fill_boxes()
Dim a As Integer
Dim applc As String
frmServiceControl.List2.Clear
frmServiceControl.List3.Clear

For a = 1 To 100
applc = GetFromINI("taskkill", "kill" + Str(a), app.Path & "\service.ini")
If applc <> "" Then frmServiceControl.List2.AddItem applc

applc = GetFromINI("taskrun", "run" + Str(a), app.Path & "\service.ini")
If applc <> "" Then frmServiceControl.List3.AddItem applc

Next a

End Function
Public Function run_tasks()
Dim xx As Integer
Dim yy As Integer
Dim appa As String
Dim found As String
Dim scan_task As String


For yy = 1 To frmServiceControl.List3.ListCount
scan_task = frmServiceControl.List3.List(yy - 1)

found = "False"
For xx = 1 To frmServiceControl.List1.ListCount



appa = frmServiceControl.List1.List(xx)
If InStr(appa, scan_task) <> 0 Then found = "True"
Next xx
If found = "False" Then
If Dir(scan_task) <> "" Then Shell scan_task, vbNormalFocus
End If

Next yy

End Function


Public Function kill_tasks()
Dim xx As Integer
Dim yy As Integer
Dim appa As String
Dim found As String
Dim scan_task As String


For yy = 1 To frmServiceControl.List2.ListCount
scan_task = frmServiceControl.List2.List(yy - 1)

found = "False"
For xx = 1 To frmServiceControl.List1.ListCount



appa = frmServiceControl.List1.List(xx - 1)
If InStr(LCase(appa), LCase(scan_task)) <> 0 Then found = appa
Next xx
If found <> "False" Then

KillTask (found)

End If

Next yy

End Function
