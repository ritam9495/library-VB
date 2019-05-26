VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form4"
   ClientHeight    =   8595
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12855
   LinkTopic       =   "Form4"
   ScaleHeight     =   8595
   ScaleWidth      =   12855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Show  Book Id:-"
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
      Left            =   2280
      MaskColor       =   &H00C0C000&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6960
      Width           =   3855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2040
      TabIndex        =   17
      Top             =   6120
      Width           =   3855
   End
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
      Height          =   735
      Left            =   11160
      MaskColor       =   &H00C0C000&
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "HOMEPAGE"
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox DataGrid1 
      Height          =   30
      Left            =   840
      ScaleHeight     =   30
      ScaleWidth      =   135
      TabIndex        =   15
      Top             =   3480
      Width           =   135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Requisition "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7800
      MaskColor       =   &H00C0C000&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   0
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Search Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   7440
      TabIndex        =   9
      Top             =   720
      Width           =   3135
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ALL"
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
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PUBLISHER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   2775
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "AUTHOR"
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
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "BOOK NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   3015
      End
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   6720
      TabIndex        =   8
      Top             =   6840
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   2160
      MaskColor       =   &H00C0C000&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5040
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   4800
      TabIndex        =   6
      Top             =   4200
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label5 
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
      Left            =   6720
      TabIndex        =   18
      Top             =   6240
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Publisher:-"
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
      Left            =   4920
      TabIndex        =   5
      Top             =   3480
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Search  Book"
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
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   5775
   End
   Begin VB.Label Label2 
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
      Height          =   735
      Left            =   2400
      TabIndex        =   2
      Top             =   3480
      Width           =   2175
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
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   3480
      Width           =   1935
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim id(3) As Integer

Dim p As Integer

Private Sub Command1_Click()
Combo1.Enabled = False
Combo1.Clear
If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" Then
MsgBox "please fill up any of the three box"
Text1.SetFocus
ElseIf Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" Then
Dim db As New ADODB.Connection
Dim rs, raz As New ADODB.Recordset
Dim sql, syl As String
id(0) = Len(Text1.Text)
id(1) = Len(Text2.Text)
id(2) = Len(Text3.Text)
p = id(0)
p = p + (10 * id(1))
p = p + (100 * id(2))
db.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\library.mdb;Persist Security Info=False")
sql = "select * from books where ID=" & p & ""
Set rs = db.Execute(sql)

If rs.BOF Or rs.EOF Then
MsgBox "The book is not available in library"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text1.SetFocus
Text4.Text = ""
Else
Text4.Text = rs(0)


Text3.Text = ""
Text2.Text = ""
Text1.Text = ""

Set rs = Nothing
db.Close
End If
ElseIf Text1.Text <> "" And Text2.Text = "" And Text3.Text = "" Then
Dim ds As New ADODB.Connection
Dim rp As New ADODB.Recordset
Dim spl As String
Command4.Enabled = False
ds.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\library.mdb;Persist Security Info=False")
spl = "select ID from books where NAME like'" & Text1.Text & "'"
Set rp = ds.Execute(spl)
If rp.BOF Or rp.EOF Then
MsgBox "The book is not available in library"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text1.SetFocus
Text4.Text = ""
Else
Text4.Text = rp(0)
Text3.Text = ""
Text2.Text = ""
Text1.Text = ""
Set rp = Nothing
ds.Close
End If
ElseIf Text1.Text = "" And Text2.Text <> "" And Text3.Text = "" Then
Dim da As New ADODB.Connection
Dim ra As New ADODB.Recordset
Dim sal As String
Combo1.Enabled = True
Command4.Enabled = True
da.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\library.mdb;Persist Security Info=False")
sal = "select NAME from books where AUTHOR like'" & Text2.Text & "'"
Set ra = da.Execute(sal)
If ra.BOF Or ra.EOF Then
MsgBox "The book is not available in library"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text2.SetFocus
Text4.Text = ""
Combo1.Enabled = False
Else
Combo1.Text = "Select the book you want"

Text3.Text = ""
Text2.Text = ""
Text1.Text = ""
Do Until ra.EOF
Combo1.AddItem (ra(0))
ra.MoveNext
Loop
Set ra = Nothing
da.Close
End If
ElseIf Text1.Text = "" And Text2.Text = "" And Text3.Text <> "" Then
Dim dz As New ADODB.Connection
Dim rz As New ADODB.Recordset
Dim szl As String
Command4.Enabled = True
dz.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\library.mdb;Persist Security Info=False")
szl = "select NAME from books where PUBLISHER like'" & Text3.Text & "'"
Set rz = dz.Execute(szl)
If rz.BOF Or rz.EOF Then
MsgBox "The book is not available in library"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text3.SetFocus
Text4.Text = ""
Combo1.Enabled = False
Else
Combo1.Enabled = True
Combo1.Text = "Select the book you want"
Text3.Text = ""
Text2.Text = ""
Text1.Text = ""
Do Until rz.EOF
Combo1.AddItem (rz(0))
rz.MoveNext
Loop
Set rz = Nothing
dz.Close
End If
End If


End Sub

Private Sub Command2_Click()
Unload Form4
Form5.Show
End Sub

Private Sub Command3_Click()
Unload Form4
Form1.Show
End Sub

Private Sub Command4_Click()
Dim da As New ADODB.Connection
Dim rb As New ADODB.Recordset
da.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\library.mdb;Persist Security Info=False")
Set rb = da.Execute("select ID from books where NAME='" & Combo1.Text & "'")
Text4.Text = rb(0)
End Sub

Private Sub Form_Activate()
Option4.Value = True
Option3.Value = False
Option1.Value = False
Option2.Value = False



End Sub

Private Sub Option1_Click()
Text1.Visible = True
Text2.Visible = False
Text3.Visible = False
End Sub

Private Sub Option2_Click()
Text1.Visible = False
Text2.Visible = True
Text3.Visible = False
End Sub

Private Sub Option3_Click()
Text1.Visible = False
Text2.Visible = False
Text3.Visible = True
End Sub

Private Sub Option4_Click()
Text1.Visible = True
Text2.Visible = True
Text3.Visible = True
End Sub

