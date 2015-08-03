VERSION 5.00
Begin VB.Form frmexam_select 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Exam"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   4140
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5520
      Top             =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C#"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   60
      TabIndex        =   3
      Top             =   3000
      Width           =   4035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SQL"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   60
      TabIndex        =   2
      Top             =   2040
      Width           =   4035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C++"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   1080
      Width           =   4035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   4035
   End
End
Attribute VB_Name = "frmexam_select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
    
    If Index = 0 Then
        entrypoint.Examtable = "c"
    ElseIf Index = 1 Then
        entrypoint.Examtable = "cpp"
    ElseIf Index = 2 Then
        entrypoint.Examtable = "oracle"
    ElseIf Index = 3 Then
        entrypoint.Examtable = "csharp"
    End If
        
    Unload Me
End Sub
 
'Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Command1(Index).BackColor = vbRed
'End Sub

'Private Sub Timer1_Timer()
    
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    'if it is admin login then no exam to be conducted
    
    
    If entrypoint.adminLogin = False Then
        Load frmexam
    Else
        entrypoint.adminLogin = False
    End If
End Sub
