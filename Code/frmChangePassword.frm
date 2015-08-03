VERSION 5.00
Begin VB.Form frmChangePassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ChangePassword"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7185
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "CANCEL"
      Height          =   495
      Index           =   1
      Left            =   5220
      TabIndex        =   7
      Top             =   2700
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Index           =   0
      Left            =   3000
      TabIndex        =   6
      Top             =   2700
      Width           =   1635
   End
   Begin VB.TextBox Text1 
      Height          =   615
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2820
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1800
      Width           =   4275
   End
   Begin VB.TextBox Text1 
      Height          =   615
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2820
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   4275
   End
   Begin VB.TextBox Text1 
      Height          =   615
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   2820
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   240
      Width           =   4275
   End
   Begin VB.Label Label1 
      Caption         =   "New Password"
      Height          =   435
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1860
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "New Password"
      Height          =   435
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1140
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Old  Password"
      Height          =   435
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
    
    If Index = 1 Then
        Unload Me
    End If
    
    Dim sql As String
    Dim rs As New ADODB.Recordset
        
    rs.Open "select * from login where username = '" & uname & "'", conn, adOpenDynamic, adLockOptimistic
    rs.MoveFirst
    
    If Index = 0 Then
        If Text1(1).Text <> Text1(2).Text Then
            MsgBox "NEW PASSWORDS DOES NOT MATCH", vbOKOnly + vbExclamation, "ERROR"
            Text1(1).SetFocus
        ElseIf Text1(1).Text = vbNullString Or Text1(2).Text = vbNullString Then
            MsgBox "ENTER NEW PASSWORD", vbOKOnly + vbExclamation, "ERROR"
            Text1(1).SetFocus
        ElseIf Text1(0).Text = rs.Fields(0) Then
             sql = "update login set password = '" & Text1(1).Text & "' where username = "
             sql = sql & "'" & uname & "'"
             conn.Execute sql
             sql = "commit"
             conn.Execute sql
             rs.Close
             MsgBox "PASSWORD UPDATED SUCCESSFULLY", vbOKOnly + vbMsgBoxRight, "SUCCESS"
        Else
            MsgBox "WRONG PASSWORD", vbOKOnly + vbExclamation, "ERROR"
            Text1(0).SetFocus
        End If
    Else
        rs.Close
    End If
   Unload Me
End Sub

