VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   7275
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13440
   FillStyle       =   3  'Vertical Line
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   13440
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "St. Thomas College of Engineering and Technology"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   3960
      Width           =   11415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome To Library "
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   0
      Top             =   720
      Width           =   5655
   End
   Begin VB.Menu mnuGoto 
      Caption         =   "Goto"
      Begin VB.Menu mnuLib 
         Caption         =   "Librarian"
      End
      Begin VB.Menu mnuStd 
         Caption         =   "Student/Teacher"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuLib_Click()
    Form2.Show
    Unload Me
End Sub

Private Sub mnuStd_Click()
    Form10.Show
    Unload Me
End Sub
