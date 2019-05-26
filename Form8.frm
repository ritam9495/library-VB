VERSION 5.00
Begin VB.Form Form8 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form8"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11670
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form8"
   ScaleHeight     =   7875
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   675
      ItemData        =   "Form8.frx":0000
      Left            =   6600
      List            =   "Form8.frx":0010
      TabIndex        =   14
      Text            =   "Select"
      Top             =   2520
      Width           =   4335
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   0
      TabIndex        =   12
      Top             =   5760
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "<="
      Height          =   615
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "HOMEPAGE"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      TabIndex        =   8
      Top             =   4320
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Submit"
      Height          =   975
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6720
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   0
      TabIndex        =   4
      Top             =   4320
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Password:-"
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   5160
      Width           =   3615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Height          =   975
      Left            =   6480
      TabIndex        =   11
      Top             =   5640
      Width           =   3975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Roll No.:-"
      Height          =   615
      Left            =   6600
      TabIndex        =   9
      Top             =   3720
      Width           =   3855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Library Id:-"
      Height          =   615
      Left            =   6480
      TabIndex        =   7
      Top             =   5040
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Batch:-"
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   3720
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Department:-"
      Height          =   615
      Left            =   6960
      TabIndex        =   3
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Student's Name:-"
      Height          =   555
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   3960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "STUDENT REGISTRATION"
      Height          =   735
      Left            =   2160
      TabIndex        =   0
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p As String

Private Sub Command1_Click()
If Text1.Text = "" Or Combo1.Text = "Select" Or Text3.Text = "" Or Text5.Text = "" Then
MsgBox "Please fill up all the boxes for registration"
Else
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sql As String
If Combo1.Text = "CSE" Then
    a = 1
ElseIf Combo1.Text = "IT" Then
    a = 2
ElseIf Combo1.Text = "EE" Then
    a = 3
Else
    a = 4
End If

p = Val(Text3.Text) * 1000 + (a * 100) + Val(Text5.Text)

db.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\library.mdb;Persist Security Info=False")
sql = "select *from Student where SID =" & p & ""
Set rs = db.Execute(sql)
If rs.BOF Or rs.EOF Then
Dim spl As String
spl = "insert into Student values(" & p & ",'" & Text1.Text & "','" & Combo1.Text & "'," & Text3.Text & ",'" & Text4.Text & "')"

db.Execute (spl)
MsgBox "Student is registered"
Text1.Text = ""
Combo1.Text = "Select"
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Label7.Caption = p
db.Close
Else
MsgBox "student is already registered, id=" & p & ""
Text1.Text = ""
Combo1.Text = "Select"
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End If
Set rs = Nothing
End If
End Sub

Private Sub Command2_Click()
Unload Form8
Form6.Show
End Sub

