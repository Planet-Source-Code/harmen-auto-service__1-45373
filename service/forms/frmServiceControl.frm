VERSION 5.00
Begin VB.Form frmServiceControl 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " NT Service installer"
   ClientHeight    =   1995
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   5040
   Icon            =   "frmServiceControl.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List3 
      Height          =   2985
      Left            =   2640
      TabIndex        =   15
      Top             =   2880
      Width           =   2175
   End
   Begin VB.ListBox List2 
      Height          =   2985
      Left            =   5040
      TabIndex        =   14
      Top             =   2880
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   2955
      IntegralHeight  =   0   'False
      Left            =   240
      TabIndex        =   13
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Install and Start"
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Foreground"
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox SERVICE_NAME 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Text            =   "Regload"
      Top             =   1200
      Width           =   3735
   End
   Begin VB.TextBox Service_File_Name 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Text            =   "C:\WINNT\notepad.exe"
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Timer tmrCheck 
      Interval        =   1000
      Left            =   6000
      Top             =   120
   End
   Begin VB.CheckBox chkSystem 
      Caption         =   "System Account"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   480
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtAccount 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Service"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "Install Service"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Service name"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Path to exe"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblAccount 
      Caption         =   "Account:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Menu Runconfig 
      Caption         =   "Controll"
      Begin VB.Menu Installrun 
         Caption         =   "Install and Start"
      End
      Begin VB.Menu Installl_service 
         Caption         =   "Install Service"
      End
      Begin VB.Menu uninstall_service 
         Caption         =   "Uninstall Servive"
      End
      Begin VB.Menu start_service 
         Caption         =   "Start Service"
      End
      Begin VB.Menu stop_service 
         Caption         =   "Stop Service"
      End
   End
   Begin VB.Menu ee 
      Caption         =   ""
   End
   Begin VB.Menu Setup 
      Caption         =   "Setup"
      Begin VB.Menu ttrun 
         Caption         =   "Tasks to run"
      End
      Begin VB.Menu ttkill 
         Caption         =   "Task to kill"
      End
   End
   Begin VB.Menu ddd 
      Caption         =   ""
   End
   Begin VB.Menu Bugs 
      Caption         =   "Bugs"
   End
End
Attribute VB_Name = "frmServiceControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(1 To 128) As Byte
End Type
Private Const VER_PLATFORM_WIN32_NT = 2&

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_SHOWNORMAL = 1&
Dim ServState As SERVICE_STATE, Installed As Boolean

Private Sub Bugs_Click()

form1.Show

End Sub

Private Sub chkSystem_Click()
    If chkSystem Then
        txtAccount = "LocalSystem"
        txtAccount.Enabled = False
        txtPassword.Enabled = False
        lblAccount.Enabled = False
        lblPassword.Enabled = False
    Else
        txtAccount = vbNullString
        txtAccount.Enabled = True
        txtPassword.Enabled = True
        lblAccount.Enabled = True
        lblPassword.Enabled = True
    End If
End Sub

Private Sub cmdInstall_Click()
    CheckService
    If Not cmdInstall.Enabled Then Exit Sub
    cmdInstall.Enabled = False
    If Installed Then
        DeleteNTService
    Else
        SetNTService
           
 
WriteToINI "Startup", "txtAccount", txtAccount, App.Path & "\service.ini"
WriteToINI "Startup", "txtPassword", txtPassword, App.Path & "\service.ini"
WriteToINI "Startup", "SERVICE_NAME", SERVICE_NAME, App.Path & "\service.ini"
WriteToINI "Startup", "Service_File_Name", Service_File_Name, App.Path & "\service.ini"
WriteToINI "service", "Fore_ground", Check1, App.Path & "\service.ini"

       
        
    End If
    CheckService
End Sub

' This sub checks service status
Private Sub CheckService()


Installl_service.Enabled = False
Installrun.Enabled = False
uninstall_service.Enabled = False
start_service.Enabled = False
stop_service.Enabled = False


    If GetServiceConfig() = 0 Then
        Installed = True
        cmdInstall.Caption = "Uninstall Service"
        uninstall_service.Enabled = True
        txtAccount.Enabled = False
        txtPassword.Enabled = False
        lblAccount.Enabled = False
        lblPassword.Enabled = False
        chkSystem.Enabled = False
      
        Service_File_Name.Enabled = False
      SERVICE_NAME.Enabled = False
       Check1.Enabled = False
       
        ServState = GetServiceStatus()
        Select Case ServState
            Case SERVICE_RUNNING
                cmdInstall.Enabled = False
                cmdStart.Caption = "Stop Service"
                uninstall_service.Enabled = False
                stop_service.Enabled = True
                
                cmdStart.Enabled = True
            Case SERVICE_STOPPED
                cmdInstall.Enabled = True
                cmdStart.Caption = "Start Service"
                start_service.Enabled = True

                cmdStart.Enabled = True
            Case Else
                cmdInstall.Enabled = False
                cmdStart.Enabled = False
        End Select
    Else
        Installed = False
        cmdInstall.Caption = "Install Service"
        Installl_service.Enabled = True
        txtAccount.Enabled = chkSystem = 0
        txtPassword.Enabled = chkSystem = 0
        lblAccount.Enabled = chkSystem = 0
        lblPassword.Enabled = chkSystem = 0
        chkSystem.Enabled = True
        cmdStart.Enabled = False
        cmdInstall.Enabled = True
        
        SERVICE_NAME.Enabled = True
        Service_File_Name.Enabled = True
        SERVICE_NAME.Enabled = True
        Check1.Enabled = True
        
        

    End If
End Sub

Private Sub cmdStart_Click()
    CheckService
    If Not cmdStart.Enabled Then Exit Sub
    cmdStart.Enabled = False
    If ServState = SERVICE_RUNNING Then
        StopNTService
    ElseIf ServState = SERVICE_STOPPED Then
        StartNTService
    End If
    CheckService
End Sub




Private Sub Command1_Click()
cmdInstall_Click
DoEvents
cmdStart_Click

End Sub









Private Sub Form_Load()
fill_boxes


Command1.Enabled = False

    If Not CheckIsNT() Then
        MsgBox "This program requires Windows NT/2000/XP"
        Unload Me
        Exit Sub
    End If
    
    
    AppPath = App.Path

    
    
    
    
    If Right$(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
    chkSystem_Click
    CheckService
    
    
txtAccount = GetFromINI("Startup", "txtAccount", App.Path & "\service.ini")
txtPassword = GetFromINI("Startup", "txtPassword", App.Path & "\service.ini")
SERVICE_NAME = GetFromINI("Startup", "SERVICE_NAME", App.Path & "\service.ini")
Service_File_Name = GetFromINI("Startup", "Service_File_Name", App.Path & "\service.ini")
Check1 = GetFromINI("service", "fore_ground", App.Path & "\service.ini")

1000  Rem
End Sub




' CheckIsNT() returns True, if the program runs
' under Windows NT or Windows 2000, and False
' otherwise.

Private Function CheckIsNT() As Boolean
    Dim OSVer As OSVERSIONINFO
    OSVer.dwOSVersionInfoSize = LenB(OSVer)
    GetVersionEx OSVer
    CheckIsNT = (OSVer.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function
Private Sub Installrun_Click()
 Command1_Click
End Sub


Private Sub Installl_service_Click()


    If Installed Then
      ''  DeleteNTService
    Else
        SetNTService
           
 
WriteToINI "Startup", "txtAccount", txtAccount, App.Path & "\service.ini"
WriteToINI "Startup", "txtPassword", txtPassword, App.Path & "\service.ini"
WriteToINI "Startup", "SERVICE_NAME", SERVICE_NAME, App.Path & "\service.ini"
WriteToINI "Startup", "Service_File_Name", Service_File_Name, App.Path & "\service.ini"
WriteToINI "service", "Fore_ground", Check1, App.Path & "\service.ini"

       
        
    End If
End Sub





Private Sub ttkill_Click()
Form2.Show
Form2.fill_kills
frmServiceControl.Hide

End Sub

Private Sub ttrun_Click()
Form2.Show
Form2.fill_runn
frmServiceControl.Hide
End Sub

Private Sub uninstall_service_Click()
    If Installed Then
        DeleteNTService
    End If
End Sub

Private Sub start_service_Click()
    If ServState = SERVICE_RUNNING Then
      ''  StopNTService
    ElseIf ServState = SERVICE_STOPPED Then
        StartNTService
    End If
End Sub

Private Sub stop_service_Click()

    If ServState = SERVICE_RUNNING Then
        StopNTService
    End If
  
End Sub

Private Sub tmrCheck_Timer()
   CheckService
    
  If cmdInstall.Caption = "Install Service" Then
  If cmdStart.Caption = "Start Service" Then
  Command1.Enabled = True
  Installrun.Enabled = True
  
  
  Else
  Command1.Enabled = False
  Installrun.Enabled = False
  End If
  Else
  Command1.Enabled = False
  Installrun.Enabled = False
  End If
    

End Sub


