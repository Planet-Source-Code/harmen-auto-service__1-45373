VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "New task"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   4215
      Begin VB.CommandButton Command4 
         Caption         =   "Insert"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Browse"
         Height          =   255
         Left            =   3120
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save and close"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete selected"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function fill_kills()
Me.Caption = "Files to Kill"
Dim a As Integer
Dim applc As String
List1.Clear


For a = 1 To 100
applc = GetFromINI("taskkill", "kill" + Str(a), App.Path & "\service.ini")
If applc <> "" Then List1.AddItem applc

Next a

End Function
Public Function fill_runn()
Me.Caption = "Files to run"
Dim a As Integer
Dim applc As String
List1.Clear


For a = 1 To 100
applc = GetFromINI("taskrun", "run" + Str(a), App.Path & "\service.ini")
If applc <> "" Then List1.AddItem applc

Next a

End Function

Private Sub Command1_Click()
If List1.ListIndex <> -1 Then List1.RemoveItem (List1.ListIndex)




End Sub

Private Sub Command2_Click()
Dim task As String

For a = 1 To 100

task = List1.List(a - 1)
If Me.Caption = "Files to run" Then WriteToINI "taskrun", "run" + Str(a), task, App.Path & "\service.ini"
If Me.Caption = "Files to Kill" Then WriteToINI "taskkill", "kill" + Str(a), task, App.Path & "\service.ini"


Next a


Me.Hide
frmServiceControl.Show
    
End Sub

Private Sub Command3_Click()
  lzFileName = OpenFile(Me.hwnd)
  Text1 = lzFileName
End Sub


Private Sub Command4_Click()
List1.AddItem Text1
End Sub
