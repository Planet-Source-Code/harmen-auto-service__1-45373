Attribute VB_Name = "Tasks"
Option Explicit
Const MAX_PATH& = 260
Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long


Private Const WM_CLOSE = &H10
Private Const PROCESS_TERMINATE As Long = &H1
Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_VM_READ = 16

Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE = &H100000
Public Const PROCESS_ALL_ACCESS = &H1F0FFF


Public Function FillProcessListNT(lstBox As ListBox) As Long

Dim cb                As Long
Dim cbNeeded          As Long
Dim NumElements       As Long
Dim ProcessIDs()      As Long
Dim cbNeeded2         As Long
Dim NumElements2      As Long
Dim Modules(1 To 200) As Long
Dim lRet              As Long
Dim ModuleName        As String
Dim nSize             As Long
Dim hProcess          As Long
Dim i                 As Long
Dim sModName          As String
Dim sChildModName     As String
Dim iModDlls          As Long
Dim iProcesses        As Integer
    
lstBox.Clear

cb = 8
cbNeeded = 96

Do While cb <= cbNeeded
    cb = cb * 2
    ReDim ProcessIDs(cb / 4) As Long
    lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
Loop

NumElements = cbNeeded / 4
    
For i = 1 To NumElements

    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
            Or PROCESS_VM_READ, 0, ProcessIDs(i))

    If hProcess Then
 
        lRet = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded2)
   
        If lRet <> 0 Then
        
            ModuleName = Space(MAX_PATH)
       
            nSize = 500
       
            lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)
         sChildModName = Str(ProcessIDs(i))
         
          sModName = left$(ModuleName, lRet)
        
111         If Len(sChildModName) < 10 Then
        sChildModName = sChildModName + " "
        GoTo 111
        End If
        
        
            lstBox.AddItem sChildModName + " " + sModName
            iProcesses = iProcesses + 1
                
            
        End If
    Else

        FillProcessListNT = 0
    End If

    lRet = CloseHandle(hProcess)
Next

FillProcessListNT = iProcesses
End Function


Public Sub KillTask(taskk)
Dim task As String
Dim ProcessID As Integer

Dim hProcess As Long

ProcessID = Val(left(taskk, 9))
task = right(taskk, Len(taskk) - 10)

    hProcess = OpenProcess(PROCESS_TERMINATE, 1, ProcessID)
     If hProcess <> 0 Then hProcess = TerminateProcess(hProcess, 0)

End Sub

