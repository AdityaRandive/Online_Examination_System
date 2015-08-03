VERSION 5.00
Begin VB.Form frmcreate 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MIT_Addmission_form"
   ClientHeight    =   8550
   ClientLeft      =   3840
   ClientTop       =   375
   ClientWidth     =   12135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   12135
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "frmcreate.frx":0000
      Left            =   7080
      List            =   "frmcreate.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   4260
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "CREATE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7380
      Width           =   2535
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "T & C"
      Height          =   315
      Left            =   780
      TabIndex        =   15
      Top             =   6720
      Width           =   6675
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   13
      Top             =   3480
      Width           =   3255
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2580
      TabIndex        =   9
      Top             =   5220
      Width           =   4695
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "frmcreate.frx":0004
      Left            =   5400
      List            =   "frmcreate.frx":002C
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4260
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "frmcreate.frx":006C
      Left            =   3600
      List            =   "frmcreate.frx":00CD
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   4260
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2580
      TabIndex        =   5
      Top             =   5820
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   2760
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   2040
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "EMAIL ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   720
      TabIndex        =   14
      Top             =   5220
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   720
      TabIndex        =   12
      Top             =   3540
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "REGISTRATION "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Index           =   7
      Left            =   240
      TabIndex        =   11
      Top             =   360
      Width           =   11595
   End
   Begin VB.Line Line3 
      BorderColor     =   &H008080FF&
      BorderWidth     =   3
      X1              =   240
      X2              =   11880
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H008080FF&
      BorderWidth     =   3
      Height          =   8295
      Left            =   120
      Top             =   120
      Width           =   11895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H008080FF&
      BorderWidth     =   3
      X1              =   240
      X2              =   11880
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H008080FF&
      BorderWidth     =   3
      Height          =   8055
      Left            =   240
      Top             =   240
      Width           =   11655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "DATE OF BIRTH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   6
      Top             =   4260
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MOBILE NO."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   660
      TabIndex        =   2
      Top             =   5940
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   2820
      Width           =   1755
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   2160
      Width           =   1455
   End
End
Attribute VB_Name = "frmcreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
    'On Error Resume Next
    
    If Text1.Text = "" Then
        MsgBox "Please Enter the Name."
        Text1.SetFocus
    ElseIf Text2.Text = "" Then
        MsgBox "Please Enter the Username."
        Text2.SetFocus
    ElseIf Text3.Text = "" Then
        MsgBox "Please Enter the Password."
        Text3.SetFocus
    ElseIf Text5.Text = "" Then
        MsgBox "Please Enter the Emailid."
        Text5.SetFocus
    ElseIf Text6.Text = "" Then
        MsgBox "Please Enter the mobile no."
        Text6.SetFocus
    ElseIf Len(Text6.Text) <> 10 Then
        MsgBox "Please Enter a valid Mobile number."
        Text3.SetFocus
    ElseIf InStr(Text5.Text, "@") = 0 Or InStr(Text5.Text, ".") = 0 Then
        MsgBox "Invalid Email ID."
        Text6.SetFocus
    ElseIf Check2.Value = False Then
        MsgBox "Please accept the Terms & Conditions"
    Else
    'Username taken or not
            rs.Open "select username from login", conn, adOpenDynamic, adLockOptimistic
            rs.MoveFirst 'position row pointer on first row of recordset
            Do While (rs.EOF = False)
               'fetch data of cols of current row
                    If rs.Fields(0) = UCase(Text2.Text) Then
                        rs.Close
                        MsgBox "UserName already Taken" & vbCrLf & "Choose another username"
                        GoTo d
                    End If
                'move to next row
                rs.MoveNext
            Loop
            rs.Close
        'all correct add user
        'Save details in login table as normalUser
        'i.e type 0
        ' insert into login values(          username,              password,    type,    sname,                                             sdob    ,                      sphone,             s_email);
        sql = "INSERT into login values(UPPER('" & Text2.Text & "'),'" & Text3.Text & "'," & CreateType & ",'" & Text1.Text & "','" & Combo2.Text & "-" & Combo3.Text & "-" & Combo1.Text & "','" & Text6.Text & "','" & Text5.Text & "')"
        conn.Execute sql
        
        sql = "commit"
        conn.Execute sql
        MsgBox "You are successfully Registered"
        Unload Me
    End If

d:  CreateType = "0"
End Sub

'regErr:
 '       MsgBox Err.Description

'End Sub

Private Sub Form_Load()
    Dim i As Single
    Dim n As Single
    Dim y As String
    
    For i = 1980 To 2013
        y = Str(i)
        n = i - 1980
        Combo1.AddItem y, n
    Next
    
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
    Combo3.ListIndex = 0
End Sub

Private Sub Text1_KeyPress(keyascii As Integer) 'name
    If Not ((keyascii >= 97 And keyascii <= 122) Or (keyascii >= 65 And keyascii <= 96) Or (keyascii = 32) Or (keyascii = 13) Or (keyascii = 8)) Then
        Text1.Text = ""
        keyascii = 0
        MsgBox "Please Enter alphabets only"
        Text1.SetFocus
    End If
End Sub

Private Sub Text3_KeyPress(keyascii As Integer) 'mobile no.
    If Not ((keyascii >= 48 And keyascii <= 58) Or (keyascii = 13) Or (keyascii = 8) Or (keyascii >= 97 And keyascii <= 122) Or (keyascii >= 65 And keyascii <= 96)) Then
        Text3.Text = ""
        MsgBox "Only alphabets or numbers allowed"
        Text3.SetFocus
    End If
End Sub

