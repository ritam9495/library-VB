VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form11"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11625
   LinkTopic       =   "Form11"
   ScaleHeight     =   7950
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "HOMEPAGE"
      Top             =   0
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "TEACHER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   7680
      TabIndex        =   5
      Top             =   2040
      Width           =   3615
      Begin VB.CommandButton Command2 
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
         Height          =   555
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3000
         Width           =   3375
      End
      Begin VB.TextBox Text4 
         Height          =   735
         Left            =   0
         TabIndex        =   14
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox Text3 
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "TEACHER ID:-"
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
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   3375
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
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "STUDENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   3735
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
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3000
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "STUDENT ID:-"
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
         Left            =   120
         TabIndex        =   8
         Top             =   1680
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
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   480
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
      Height          =   1695
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3255
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Student"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Teacher"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "RETURN BOOKS"
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
      Left            =   3720
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l, z As Date
Dim p, q As Integer

Private Sub Command1_Click()


If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "Please fill up all boxes"
Else
Dim db As New ADODB.Connection
Dim rs, rb, rv As New ADODB.Recordset
Dim sql, spl, swl As String
db.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\library.mdb;Persist Security Info=False")

sql = "select *from books where ID=" & Val(Text1.Text) & ""
spl = "select * from student where SID=" & Val(Text2.Text) & ""
Set rs = db.Execute(sql)
Set rb = db.Execute(spl)
If rs.BOF Or rs.EOF Then
MsgBox "The BookID is invalid"
ElseIf rb.BOF Or rb.EOF Then
MsgBox "StudentID is invalid"
Else
swl = "select *from Reqs where ID=" & Val(Text1.Text) & "and SID=" & Val(Text2.Text) & ""
Set rv = db.Execute(swl)
If rv.BOF Or rv.EOF Then
MsgBox "There is no such requisition "
Set rv = Nothing
db.Close
Else
Dim rf, rz, rk As New ADODB.Recordset
Dim sfl, sol, skl As String
q = rs(4) + 1
skl = "select *from Reqs where (ID=" & Val(Text1.Text) & ")and (SID=" & Val(Text2.Text) & ")"
Set rk = db.Execute(skl)
If l > rk(3) Then
z = rk(3)
 l = l - z
MsgBox "You have to give  " & l & "/-for late submission"
l = DateValue(Now)
sfl = "delete from reqs where ID=" & Val(Text1.Text) & "and SID=" & Val(Text2.Text) & ""
Set rf = db.Execute(sfl)
sol = "update books set COPY='" & q & "'where ID = " & Text1.Text & ""
Set rz = db.Execute(sol)
Text1.Text = ""
Text2.Text = ""
Else
Dim sbl, sil As String
sbl = "delete from reqs where ID=" & Val(Text1.Text) & "and SID=" & Val(Text2.Text) & ""
Set rf = db.Execute(sbl)
sil = "update books set COPY='" & q & "'where ID = " & Text1.Text & ""
Set rz = db.Execute(sil)
MsgBox "Submission is successful"
Text1.Text = ""
Text2.Text = ""
End If
Set rk = Nothing
Set rz = Nothing
Set rs = Nothing
Set rb = Nothing
Set rf = Nothing
db.Close
End If
End If
End If
End Sub

Private Sub Command3_Click()
Unload Form11
Form6.Show
End Sub

Private Sub Command2_Click()


If Text3.Text = "" Or Text4.Text = "" Then
MsgBox "Please fill up all boxes"
Else
Dim db As New ADODB.Connection
Dim rs, rb, rv As New ADODB.Recordset
Dim sql, spl, swl As String
db.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\library.mdb;Persist Security Info=False")

sql = "select *from books where ID=" & Val(Text3.Text) & ""
spl = "select * from teacher where TID=" & Val(Text4.Text) & ""
Set rs = db.Execute(sql)
Set rb = db.Execute(spl)
If rs.BOF Or rs.EOF Then
MsgBox "The BookID is invalid"
ElseIf rb.BOF Or rb.EOF Then
MsgBox "TeacherID is invalid"
Else
swl = "select *from Reqt where ID=" & Val(Text3.Text) & "and TID=" & Val(Text4.Text) & ""
Set rv = db.Execute(swl)
If rv.BOF Or rv.EOF Then
MsgBox "There is no such requisition "
Set rv = Nothing
db.Close
Else
Dim rf, rz, rk As New ADODB.Recordset
Dim sfl, sol, skl As String
q = rs(4) + 1
skl = "select *from Reqt where (ID=" & Val(Text3.Text) & ")and (TID=" & Val(Text4.Text) & ")"
Set rk = db.Execute(skl)
If l > rk(3) Then
z = rk(3)
l = l - z
MsgBox "You have to give  " & l & "/- for late submission"
l = DateValue(Now)
sfl = "delete from Reqt where ID=" & Val(Text3.Text) & "and TID=" & Val(Text4.Text) & ""
Set rf = db.Execute(sfl)
sol = "update books set COPY='" & q & "'where ID = " & Val(Text3.Text) & ""
Set rz = db.Execute(sol)
Text3.Text = ""
Text4.Text = ""
Else
Dim sbl, sil As String
sbl = "delete from reqt where ID=" & Val(Text3.Text) & "and TID=" & Val(Text4.Text) & ""
Set rf = db.Execute(sbl)
sil = "update books set COPY='" & q & "'where ID = " & Text3.Text & ""
Set rz = db.Execute(sil)
MsgBox "Submission is successful"
Text3.Text = ""
Text4.Text = ""
End If
Set rk = Nothing
Set rz = Nothing
Set rs = Nothing
Set rb = Nothing
Set rf = Nothing
db.Close
End If
End If
End If
End Sub

Private Sub Form_Activate()
Frame2.Visible = False
Frame3.Visible = False
Option1.Value = False
Option2.Value = False
l = DateValue(Now)
End Sub

Private Sub Option1_Click()
Frame3.Visible = True
Frame2.Visible = False
End Sub

Private Sub Option2_Click()
Frame3.Visible = False
Frame2.Visible = True
End Sub

