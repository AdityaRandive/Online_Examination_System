VERSION 5.00
Begin VB.Form frmAdmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administration"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14565
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
   ScaleHeight     =   9360
   ScaleWidth      =   14565
   Begin VB.CommandButton Command4 
      Caption         =   "FULL REPORT"
      Height          =   1335
      Left            =   4500
      TabIndex        =   16
      Top             =   7560
      Width           =   5535
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   14055
      Begin VB.Frame Frame4 
         Caption         =   "QUESTIONS"
         Height          =   6135
         Left            =   9900
         TabIndex        =   9
         Top             =   540
         Width           =   3195
         Begin VB.CommandButton Command3 
            Caption         =   "VIEW"
            Height          =   1035
            Index           =   3
            Left            =   300
            TabIndex        =   15
            Top             =   4860
            Width           =   2595
         End
         Begin VB.CommandButton Command3 
            Caption         =   "MODIFY"
            Height          =   1035
            Index           =   2
            Left            =   300
            TabIndex        =   12
            Top             =   3540
            Width           =   2595
         End
         Begin VB.CommandButton Command3 
            Caption         =   "DELETE"
            Height          =   1035
            Index           =   1
            Left            =   300
            TabIndex        =   11
            Top             =   2220
            Width           =   2595
         End
         Begin VB.CommandButton Command3 
            Caption         =   "ADD"
            Height          =   1035
            Index           =   0
            Left            =   300
            TabIndex        =   10
            Top             =   900
            Width           =   2595
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "STUDENT"
         Height          =   6135
         Left            =   480
         TabIndex        =   5
         Top             =   600
         Width           =   3195
         Begin VB.CommandButton Command1 
            Caption         =   "VIEW"
            Height          =   1035
            Index           =   3
            Left            =   240
            TabIndex        =   13
            Top             =   4860
            Width           =   2595
         End
         Begin VB.CommandButton Command1 
            Caption         =   "MODIFY"
            Height          =   1035
            Index           =   2
            Left            =   240
            TabIndex        =   8
            Top             =   3540
            Width           =   2595
         End
         Begin VB.CommandButton Command1 
            Caption         =   "DELETE"
            Height          =   1035
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Top             =   2220
            Width           =   2595
         End
         Begin VB.CommandButton Command1 
            Caption         =   "ADD"
            Height          =   1035
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   900
            Width           =   2595
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "ADMINISTRATOR"
         Height          =   6075
         Left            =   5160
         TabIndex        =   1
         Top             =   600
         Width           =   3195
         Begin VB.CommandButton Command2 
            Caption         =   "VIEW"
            Height          =   1035
            Index           =   3
            Left            =   300
            TabIndex        =   14
            Top             =   4860
            Width           =   2595
         End
         Begin VB.CommandButton Command2 
            Caption         =   "ADD"
            Height          =   1035
            Index           =   0
            Left            =   300
            TabIndex        =   4
            Top             =   900
            Width           =   2595
         End
         Begin VB.CommandButton Command2 
            Caption         =   "DELETE"
            Height          =   1035
            Index           =   1
            Left            =   300
            TabIndex        =   3
            Top             =   2220
            Width           =   2595
         End
         Begin VB.CommandButton Command2 
            Caption         =   "MODIFY"
            Height          =   1035
            Index           =   2
            Left            =   300
            TabIndex        =   2
            Top             =   3540
            Width           =   2595
         End
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset

Private Sub Command1_Click(Index As Integer)
    Me.Hide
    
    Select Case Index
        Case 0: Load frmcreate
                
        Case 1 To 3:
                Load frmAdmiStudent
    End Select
    
End Sub

Private Sub Command2_Click(Index As Integer)

        
    Select Case Index
        Case 0:    'Open table selection form by administrative privilages
                CreateType = "3"
                entrypoint.adminLogin = True
                Me.Hide
                
        Case 1 To 3:
                Load frmAdmiAdMINISTRATOR
    End Select
End Sub

Private Sub Command3_Click(Index As Integer)    'Question
    Load frmAdmiQuestions
End Sub

Private Sub Command4_Click()
    Load SHORT_REPORT
    SHORT_REPORT.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    frmAdmin.Show
End Sub
