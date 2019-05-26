VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form9"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8730
   LinkTopic       =   "Form9"
   ScaleHeight     =   6225
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form9.frx":0000
      Left            =   120
      List            =   "Form9.frx":0010
      TabIndex        =   12
      Text            =   "Select your Designation"
      Top             =   4080
      Width           =   4095
   End
   Begin VB.TextBox Text4 
      Height          =   855
      Left            =   4440
      TabIndex        =   10
      Top             =   3720
      Width           =   4215
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
      Height          =   615
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "HOMEPAGE"
      Top             =   0
      Width           =   1935
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
      Height          =   615
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   4560
      TabIndex        =   4
      Top             =   1800
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   4335
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
      Height          =   615
      Left            =   4560
      TabIndex        =   11
      Top             =   3000
      Width           =   3975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
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
      Left            =   0
      TabIndex        =   9
      Top             =   5400
      Width           =   3495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Library Id:-"
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
      Left            =   0
      TabIndex        =   7
      Top             =   4800
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Designation:-"
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
      TabIndex        =   5
      Top             =   3000
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Department:-"
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
      Left            =   4560
      TabIndex        =   3
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Teacher's Name:-"
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
      TabIndex        =   1
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Teacher Registration"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p, id(3) As Integer

Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
MsgBox "Please fill up all the boxes for registration"
Else
id(0) = Len(Text1.Text)
id(1) = Len(Text2.Text)
id(2) = Len(Text3.Text)
p = id(0) + (10 * id(1)) + (100 * id(2))
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sql As String
db.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\library.mdb;Persist Security Info=False")
sql = "select *from Teacher where TID=" & p & ""
Set rs = db.Execute(sql)
If (rs.BOF Or rs.EOF) Then
Dim spl As String
srl = "insert into Teacher values(" & p & ",'" & Text1.Text & "','" & Text2.Text & "','" & Text4.Text & "')"
db.Execute (srl)
MsgBox "Registration is completed"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Label6.Caption = p
Else
MsgBox "Teacher is aready registered, id=" & rs(0) & ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End If
db.Close
Set rs = Nothing
End If

End Sub

Private Sub Command2_Click()
Unload Form9
Form6.Show

End Sub
