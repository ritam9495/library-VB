VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form6"
   ClientHeight    =   5115
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11505
   LinkTopic       =   "Form6"
   ScaleHeight     =   5115
   ScaleWidth      =   11505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Return"
      Height          =   1095
      Left            =   7920
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Member Registraton"
      Height          =   1095
      Left            =   4200
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Book"
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Menu mnuLO 
      Caption         =   "Log Out"
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form3.Show
    Unload Me
End Sub

Private Sub Command2_Click()
    Form7.Show
    Unload Me
End Sub

Private Sub Command3_Click()
    Form11.Show
    Unload Me
End Sub

Private Sub mnuLO_Click()
    Form1.Show
    Unload Me
End Sub
