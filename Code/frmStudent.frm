VERSION 5.00
Begin VB.Form frmStudent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Administration"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11100
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   15.75
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
   ScaleHeight     =   7155
   ScaleWidth      =   11100
   Begin VB.CommandButton Command1 
      Caption         =   "GET FULL REPORT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   4
      Left            =   1800
      TabIndex        =   6
      Top             =   5400
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGOUT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Index           =   3
      Left            =   6360
      TabIndex        =   5
      Top             =   5340
      Width           =   2955
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GIVE EXAM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   0
      Left            =   4200
      TabIndex        =   2
      Top             =   1140
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CHANGE PASSWORD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   3780
      Width           =   3435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DELETE ACCOUNT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   2
      Left            =   6360
      TabIndex        =   0
      Top             =   3780
      Width           =   2955
   End
   Begin VB.Label Label1 
      Caption         =   "Ready To Test Your Knowledge - Give Exam"
      Height          =   615
      Left            =   2520
      TabIndex        =   4
      Top             =   420
      Width           =   6315
   End
   Begin VB.Label Label2 
      Caption         =   "Want to Change Password Or Delete Account"
      Height          =   555
      Left            =   2280
      TabIndex        =   3
      Top             =   3000
      Width           =   6915
   End
End
Attribute VB_Name = "frmStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
       
    Select Case Index
        
        Case 0  'Give Exam
                MDIadmin.Hide
                Load frmexam_select
                frmexam_select.Show
        
        Case 1  'Chang Password
                Load frmChangePassword
                frmChangePassword.Show
        
        Case 2  'Delete account
                DeleteAccount
                
        Case 3  'LogOut
                'conn.Execute "DELETE FROM report where rid = '" & uname & "'"
                'conn.Execute "commit"
                conn.Execute "delete from report where rid = '" & uname & "'"
                
                uname = "---UNKNOWN---"
                Unload MDIadmin
        Case 4: 'Student Report
                Load frmView
    End Select
End Sub

