VERSION 5.00
Begin VB.UserControl Timer 
   BackColor       =   &H00FFFF00&
   ClientHeight    =   2925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8355
   ScaleHeight     =   2925
   ScaleWidth      =   8355
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.Timer Timer1 
         Left            =   840
         Top             =   1560
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remaining Time :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   600
         TabIndex        =   1
         Top             =   120
         Width           =   3885
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   60
         Picture         =   "Timer.ctx":0000
         Top             =   60
         Width           =   480
      End
   End
End
Attribute VB_Name = "Timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim min As Single
Dim sec As Single
Dim flag As Boolean

Private Sub Timer1_Timer()
    Label1.Caption = "Remaining Time : " & min & ":" & sec
    sec = sec - 1
        
    If flag = False Then
      Label1.Caption = "Remaining Time : " & min & ":" & sec
      If sec = 0 Then
          min = min - 1
      Label1.Caption = "Remaining Time : " & min & ":" & sec
          sec = 59
      Label1.Caption = "Remaining Time : " & min & ":" & sec
      End If
    Else
        If sec = 0 Then
            min = 0
            sec = 0
            Label1.Caption = "Remaining Time : " & min & ":" & sec
            Timer1.Enabled = False
        End If
        Label1.Caption = "Remaining Time : " & min & ":" & sec
    End If
    
    
    If min = 0 And flag = False Then
            Label1.Caption = "Remaining Time : " & min & ":" & sec
            min = 0
            flag = True
    End If
    
    If min = 9 Then
            Frame1.BackColor = vbRed
    End If
    
End Sub

Private Sub UserControl_Initialize()
    min = 29    'ExamTime -1
    sec = 59
    Label1.Caption = "Remaining Time : " & min & ":" & sec
    flag = False
        
    Timer1.Interval = 1000
    Timer1.Enabled = True
End Sub
