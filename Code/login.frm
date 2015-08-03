VERSION 5.00
Begin VB.Form frmlogin 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "LOGIN"
   ClientHeight    =   11385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   22215
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7.906
   ScaleMode       =   5  'Inch
   ScaleWidth      =   15.427
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton abt 
      Caption         =   "ABOUT"
      Height          =   735
      Left            =   7200
      TabIndex        =   9
      Top             =   8280
      Width           =   2235
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      Height          =   735
      Left            =   10020
      TabIndex        =   8
      Top             =   8280
      Width           =   2235
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "AUTHENTICATION"
      Height          =   3135
      Left            =   7380
      TabIndex        =   1
      Top             =   3780
      Width           =   4635
      Begin VB.CommandButton Command1 
         Caption         =   "Create Account"
         Height          =   450
         Index           =   1
         Left            =   2040
         TabIndex        =   7
         Top             =   2400
         Width           =   2355
      End
      Begin VB.TextBox Text1 
         Height          =   450
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1560
         Width           =   2835
      End
      Begin VB.TextBox Text1 
         Height          =   450
         Index           =   0
         Left            =   1560
         TabIndex        =   3
         Top             =   885
         Width           =   2835
      End
      Begin VB.CommandButton Command1 
         Caption         =   " Login"
         Height          =   450
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   2400
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Password"
         Height          =   555
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1575
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Username"
         Height          =   555
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   900
         Width           =   1335
      End
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Height          =   1095
      Left            =   6900
      Top             =   8100
      Width           =   5595
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Height          =   1335
      Left            =   6780
      Top             =   7980
      Width           =   5835
   End
   Begin VB.Shape Shape5 
      BorderWidth     =   2
      Height          =   3255
      Left            =   7260
      Top             =   3720
      Width           =   4815
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Height          =   4215
      Left            =   6780
      Top             =   3300
      Width           =   5835
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Height          =   3975
      Left            =   6900
      Top             =   3420
      Width           =   5595
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   1335
      Left            =   5280
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   9315
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   1095
      Left            =   5400
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   9075
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ONLINE EXAMINATION"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   1380
      Width           =   18135
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Recordset

Private Sub Command1_Click(Index As Integer)   'login
   
    'to be secured from deleted account
    If Text1(1).Text = vbNullString And Index = 0 Then
        MsgBox "Wrong Username or Password", vbCritical + vbOKOnly
        Text1(1).SetFocus
        GoTo e
    End If
        
    loginOK = False 'fro n attempts
    entrypoint.CreateAccount = False
           
    If Index = 0 Then   'User
        loginverification Text1(0).Text, Text1(1).Text
        If loginOK = True Then
            MDIadmin.Show
            Unload Me
        End If
    ElseIf Index = 1 Then   'Create New User
        entrypoint.CreateAccount = True
        MDIadmin.Show
        Unload Me
    End If
        
e:
End Sub

Private Sub Command2_Click()
    entrypoint.exitpoint
End Sub

Private Sub Form_Load()
    loginOK = False
End Sub

Private Sub Text1_KeyPress(Index As Integer, keyascii As Integer)
    If keyascii = 13 Then
        Command1_Click (0)  'On enter key press call login command
    End If
End Sub
