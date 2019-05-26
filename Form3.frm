VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11175
   LinkTopic       =   "Form3"
   ScaleHeight     =   5040
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   975
      Left            =   2760
      TabIndex        =   11
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "<="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   7800
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   3075
      TabIndex        =   8
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   975
      Left            =   5160
      TabIndex        =   5
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   2760
      TabIndex        =   4
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Copy:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Publisher:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   6
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Author:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2760
      TabIndex        =   3
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Enlist Book In Library"
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
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Name:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p, c, d As Integer
Dim id(3) As Integer

Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "Please enter book's name"
Text1.SetFocus
ElseIf Text2.Text = "" Then
MsgBox "Please enter author's name"
Text2.SetFocus
ElseIf Text3.Text = "" Then
MsgBox "Please enter publisher's name"
Text3.SetFocus
ElseIf Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" Then
id(0) = Len(Text1.Text)
id(1) = Len(Text2.Text)
id(2) = Len(Text3.Text)
p = id(0)
p = p + (10 * id(1))
p = p + (100 * id(2))
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sql, srl, spl As String
db.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\library.mdb;Persist Security Info=False")
sql = "select *from books where id=" & p & ""
Set rs = db.Execute(sql)
    If rs.BOF Or rs.EOF Then
srl = "insert into books values (" & p & ",'" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "'," & Text4.Text & ")"
db.Execute (srl)
MsgBox "Enlistment is complete"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
db.Close
Else
d = rs(4)
srl = "DELETE FROM books WHERE ID=" & p & ""
db.Execute (srl)


spl = "insert into books values(" & p & ",'" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "'," & d + 1 & ")"
db.Execute (spl)
MsgBox "COPY number is increased,Enlistment is complete"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
db.Close

End If
Set rs = Nothing
End If


End Sub

Private Sub Command2_Click()
Unload Form3
Form6.Show
End Sub

Private Sub Form_Activate()
c = 1
End Sub

