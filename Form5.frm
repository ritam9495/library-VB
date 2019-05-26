VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form5"
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12810
   LinkTopic       =   "Form5"
   ScaleHeight     =   8415
   ScaleWidth      =   12810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
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
      Height          =   615
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "HOMEPAGE"
      Top             =   120
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "TEACHER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   4095
      Begin VB.TextBox Text5 
         Height          =   615
         Left            =   120
         TabIndex        =   18
         Top             =   4080
         Width           =   3735
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4920
         Width           =   3615
      End
      Begin VB.TextBox Text4 
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   3735
      End
      Begin VB.TextBox Text3 
         Height          =   735
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Password:-"
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
         Left            =   0
         TabIndex        =   19
         Top             =   3360
         Width           =   4095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "TEACHER  ID:-"
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
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   4095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "BOOK ID:-"
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
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "STUDENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   4200
      TabIndex        =   5
      Top             =   1800
      Width           =   3975
      Begin VB.TextBox Text6 
         Height          =   735
         Left            =   120
         TabIndex        =   20
         Top             =   4080
         Width           =   3495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "SUBMIT"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5040
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Height          =   855
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Password:-"
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
         Left            =   120
         TabIndex        =   21
         Top             =   3480
         Width           =   3495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Student ID:-"
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
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "BOOK ID:-"
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
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Teacher/Student"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5655
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Student"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   480
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Teacher"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "TO GET THE ID OF DESIRED BOOK"
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "SUBMIT YOUR REQUISITION"
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
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As Date
Dim s, i As String
Dim p, q As Integer

Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "Please fill up all boxes"
Else
Dim db As New ADODB.Connection
Dim rs, rb As New ADODB.Recordset
Dim sql, spl As String
db.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\library.mdb;Persist Security Info=False")
sql = "select *from books where ID=" & Val(Text1.Text) & ""
spl = "select * from student where SID=" & Val(Text2.Text) & ""
Set rs = db.Execute(sql)
Set rb = db.Execute(spl)

If rs.BOF Or rs.EOF Then
MsgBox "The BookID is invalid"
ElseIf rb.BOF Or rb.EOF Then
MsgBox "StudentID is invalid"
ElseIf rs(4) <> 0 And rb(4) = Text6.Text Then
Dim rf, rz As New ADODB.Recordset
Dim sfl, sol As String
p = rs(4)
q = p - 1

sfl = "insert into reqs values(" & rs(0) & "," & rb(0) & ",#" & t & "#,#" & t + 14 & "#)"
Set rf = db.Execute(sfl)
MsgBox "Requistion is completed, date of return is " & t + 14 & ""

sol = "update books set COPY='" & q & "'where ID = " & Text1.Text & ""
Set rz = db.Execute(sol)
Text1.Text = ""
Text2.Text = ""
Text6.Text = ""
ElseIf rs(4) = 0 Then
MsgBox "No copy of this book is avalaible now,come later"
Text1.Text = ""
Text2.Text = ""
Text6.Text = ""
Else
MsgBox "Password is incorrect"
End If
Set rz = Nothing
Set rs = Nothing
Set rb = Nothing
Set rf = Nothing
db.Close
End If
End Sub

Private Sub Command2_Click()
Unload Form5
Form4.Show
End Sub

Private Sub Command3_Click()
If Text3.Text = "" Or Text4.Text = "" Then
MsgBox "Please fill up all boxes"
Else
Dim db As New ADODB.Connection
Dim rs, rb As New ADODB.Recordset
Dim sql, spl As String
db.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\library.mdb;Persist Security Info=False")
sql = "select *from books where ID=" & Val(Text3.Text) & ""
spl = "select * from teacher where TID=" & Val(Text4.Text) & ""
Set rs = db.Execute(sql)
Set rb = db.Execute(spl)
If rs.BOF Or rs.EOF Then
MsgBox "The BookID is invalid"
ElseIf rb.BOF Or rb.EOF Then
MsgBox "TeacherID is invalid"
ElseIf rs(4) <> 0 And rb(4) = Text5.Text Then
Dim rf, rz As New ADODB.Recordset
Dim sfl, sol As String

sfl = "insert into reqt values(" & rs(0) & "," & rb(0) & ",#" & t & "#,#" & t + 14 & "#)"
Set rf = db.Execute(sfl)
MsgBox "Requistion is completed, date of return is " & t + 14 & ""
p = rs(4)
q = p - 1
sol = "update books set COPY='" & q & "'where ID = " & Text3.Text & ""
Set rz = db.Execute(sol)
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf rs(4) = 0 Then
MsgBox "No copy of this book is avalaible now,come later"
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Else
    MsgBox "Password is incorrect"
End If
Set rz = Nothing
Set rs = Nothing
Set rb = Nothing
Set rf = Nothing
db.Close
End If
End Sub

Private Sub Command4_Click()
Unload Form5
Form1.Show


End Sub

Private Sub Form_Activate()
Frame2.Visible = False
Frame3.Visible = False
t = DateValue(Now)
End Sub

Private Sub Option1_Click()
Frame3.Visible = True
Frame2.Visible = False
End Sub

Private Sub Option2_Click()
Frame3.Visible = False
Frame2.Visible = True
End Sub

