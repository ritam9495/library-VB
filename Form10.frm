VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form10"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11715
   LinkTopic       =   "Form10"
   ScaleHeight     =   4560
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Search"
      Height          =   1455
      Left            =   5880
      TabIndex        =   1
      Top             =   1200
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Requisition"
      Height          =   1455
      Left            =   1800
      TabIndex        =   0
      Top             =   1200
      Width           =   2895
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form5.Show
    Unload Me
End Sub

Private Sub Command2_Click()
    Form4.Show
    Unload Me
End Sub
