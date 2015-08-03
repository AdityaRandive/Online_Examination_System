VERSION 5.00
Begin VB.Form frmResult 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RESULT"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10920
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10875
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   5
         Left            =   6540
         TabIndex        =   7
         Top             =   1800
         Width           =   3615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "TIME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   4
         Left            =   480
         TabIndex        =   6
         Top             =   1800
         Width           =   4335
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFF00&
         Height          =   5415
         Index           =   1
         Left            =   180
         Shape           =   4  'Rounded Rectangle
         Top             =   2220
         Width           =   10455
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFF00&
         Height          =   5535
         Index           =   0
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   2160
         Width           =   10575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "YOU ARE PASSED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   795
         Index           =   3
         Left            =   1200
         TabIndex        =   5
         Top             =   6300
         Width           =   8295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "YOUR SCORE IS   :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   795
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   4440
         Width           =   7035
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SCORE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   795
         Index           =   2
         Left            =   7380
         TabIndex        =   3
         Top             =   4440
         Width           =   2715
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "USER NAME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   795
         Index           =   0
         Left            =   420
         TabIndex        =   2
         Top             =   2640
         Width           =   9675
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RESULT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1575
         Left            =   2460
         TabIndex        =   1
         Top             =   120
         Width           =   5775
      End
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub Form_Load()
    
    Label2(4).Caption = time
    Label2(5).Caption = Date
    
    rs.Open "select * from report where rid = '" & uname & "'", conn, adOpenDynamic, adLockOptimistic
    rs.MoveFirst
    
    Label2(0).Caption = UCase(uname)
    Label2(2).Caption = rs.Fields(1)
    
    If rs.Fields(1) > 12 Then
        Label2(3).Caption = "YOU ARE PASSED"
    Else
        Label2(3).Caption = "YOU ARE FAILED"
    End If
    Dim t As String
    Dim d As String
    Dim sql As String
    
    t = time
    d = Date
    
    sql = "insert into fullreport values('" & uname & "','" & Examtable & "','" & rs.Fields(1) & "','" & d & "','" & t & "')"
    
    conn.Execute sql
    
    rs.Close
End Sub



Private Sub Form_Unload(Cancel As Integer)
 Unload frmexam
End Sub

