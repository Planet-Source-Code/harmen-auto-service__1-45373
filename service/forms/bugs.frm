VERSION 5.00
Begin VB.Form form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bugs"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5130
   ClipControls    =   0   'False
   Icon            =   "bugs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   $"bugs.frx":0442
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label1 = ""

add_text ("In the late 40's even a simple computer was a big thing:")
add_text ("1000's of vacuum tubes and 1000's of square feet of floor space.")
add_text ("A group of programmers were working late one hot summer night.")
add_text ("To help to dissipate all the heat generated by those tubes, all ")
add_text ("the windows were open. At one point the program that they")
add_text ("were working on bombed-out. Eventually they found the problem:")
add_text ("a moth had flown in and had become lodged in the wiring, creating ")
add_text ("a short-circuit. Afterwards, every time a program would crash the ")
add_text ("programmer would exclaim, There must be a bug in the machine!")
add_text ("To this day that has remained one of the mainstays of programmers:")
add_text ("")

add_text ("When the program goes wrong, blame the hardware!")


End Sub

Private Sub add_text(text)
Label1.Caption = Label1.Caption + text + Chr(13)

End Sub

