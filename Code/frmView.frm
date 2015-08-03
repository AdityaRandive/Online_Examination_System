VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmView 
   AutoRedraw      =   -1  'True
   Caption         =   "View"
   ClientHeight    =   8655
   ClientLeft      =   165
   ClientTop       =   495
   ClientWidth     =   9315
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmView.frx":0000
   ScaleHeight     =   8655
   ScaleWidth      =   9315
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4935
      Left            =   120
      TabIndex        =   1
      Top             =   3060
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   8705
      _Version        =   393217
      TextRTF         =   $"frmView.frx":300C42
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   300
      TabIndex        =   0
      Top             =   540
      Width           =   9975
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub Form_Load()
    Dim sql As String
    
    Label1.Caption = "ONLINE EXAM REPORT OF : " & uname
    sql = "select * from fullreport where username = '" & uname & "'"
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic
    
    If rs.EOF Then
        MsgBox "No tests given yet"
        GoTo e
    End If
    rs.MoveFirst

    Dim n As Single
    
    RichTextBox1.Text = "no. )     USERNAME     TESTNAME            MARKS        TEST DATE       TEST TIME" & vbCrLf
    Do Until rs.EOF
        n = 1
        RichTextBox1.Text = RichTextBox1.Text & n & " )     " & rs.Fields(0) & "     " & rs.Fields(1) & "            " & rs.Fields(2) & "        " & rs.Fields(3) & "       " & rs.Fields(4) & vbCrLf
        rs.MoveNext
        n = n + 1
    Loop
e:
    rs.Close
End Sub

