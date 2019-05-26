VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form7"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11370
   LinkTopic       =   "Form7"
   ScaleHeight     =   7245
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "TEACHER REGISTRATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   6495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "STUDENT REGISTRATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   6135
   End
   Begin VB.Label OR 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "OR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   2880
      Width           =   2175
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form8.Show
Unload Form7
End Sub

Private Sub Command2_Click()
Form9.Show
Unload Form7

End Sub
