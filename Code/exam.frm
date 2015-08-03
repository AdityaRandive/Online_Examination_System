VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "*\ATimer\Timer.vbp"
Begin VB.Form frmexam 
   BorderStyle     =   0  'None
   Caption         =   "x"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
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
   Icon            =   "exam.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   2  'Custom
   Picture         =   "exam.frx":0442
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ActiveXTimer.Timer Timer5 
      Height          =   555
      Left            =   16380
      TabIndex        =   49
      Top             =   900
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   979
   End
   Begin RichTextLib.RichTextBox question 
      DataSource      =   "DataEnvironment1"
      Height          =   3255
      Left            =   420
      TabIndex        =   48
      Top             =   2460
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   5741
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"exam.frx":301084
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   18780
      Top             =   1020
   End
   Begin VB.PictureBox Timer3 
      Height          =   555
      Left            =   16380
      ScaleHeight     =   495
      ScaleWidth      =   3915
      TabIndex        =   44
      Top             =   900
      Width           =   3975
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   7815
      Left            =   15600
      TabIndex        =   7
      Top             =   2460
      Width           =   4575
      Begin VB.Timer Timer2 
         Left            =   180
         Top             =   240
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000FF&
         Caption         =   "SUBMIT"
         Height          =   615
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Get your Score"
         Top             =   6960
         Width           =   3495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "30"
         Height          =   555
         Index           =   29
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   5160
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "29"
         Height          =   555
         Index           =   28
         Left            =   2700
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   5160
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "28"
         Height          =   555
         Index           =   27
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   5160
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "27"
         Height          =   555
         Index           =   26
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   5160
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "26"
         Height          =   555
         Index           =   25
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   5160
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "25"
         Height          =   555
         Index           =   24
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   4380
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "24"
         Height          =   555
         Index           =   23
         Left            =   2700
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   4380
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "23"
         Height          =   555
         Index           =   22
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   4380
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "22"
         Height          =   555
         Index           =   21
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   4380
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "21"
         Height          =   555
         Index           =   20
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   4380
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "20"
         Height          =   555
         Index           =   19
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "19"
         Height          =   555
         Index           =   18
         Left            =   2700
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "18"
         Height          =   555
         Index           =   17
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "17"
         Height          =   555
         Index           =   16
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "16"
         Height          =   555
         Index           =   15
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "15"
         Height          =   555
         Index           =   14
         Left            =   3480
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2820
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "14"
         Height          =   555
         Index           =   13
         Left            =   2700
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2820
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "13"
         Height          =   555
         Index           =   12
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2820
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "12"
         Height          =   555
         Index           =   11
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2820
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "11"
         Height          =   555
         Index           =   10
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2820
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "10"
         Height          =   555
         Index           =   9
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "9"
         Height          =   555
         Index           =   8
         Left            =   2700
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "8"
         Height          =   555
         Index           =   7
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "7"
         Height          =   555
         Index           =   6
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "6"
         Height          =   555
         Index           =   5
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "5"
         Height          =   555
         Index           =   4
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1260
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "4"
         Height          =   555
         Index           =   3
         Left            =   2700
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1260
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "3"
         Height          =   555
         Index           =   2
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1260
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "2"
         Height          =   555
         Index           =   1
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1260
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "1"
         Height          =   555
         Index           =   0
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1260
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1380
         TabIndex        =   43
         ToolTipText     =   "No. of questions attempted"
         Top             =   6120
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "All Questions"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1140
         TabIndex        =   9
         Top             =   180
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   16440
      TabIndex        =   5
      Top             =   120
      Width           =   3915
      Begin VB.Image Image3 
         Height          =   690
         Left            =   0
         Picture         =   "exam.frx":301106
         Stretch         =   -1  'True
         Top             =   0
         Width           =   750
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Height          =   675
         Left            =   780
         TabIndex        =   6
         ToolTipText     =   "username"
         Top             =   60
         Width           =   3075
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "PREV.   QUESION"
      Enabled         =   0   'False
      Height          =   615
      Index           =   1
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Go to previous Question"
      Top             =   10560
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "NEXT QUESION"
      Height          =   615
      Index           =   0
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Go to next Question"
      Top             =   10560
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "EXIT"
      Height          =   615
      Left            =   16080
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Exit Exam"
      Top             =   10620
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Options"
      Height          =   4395
      Left            =   420
      TabIndex        =   0
      Top             =   5880
      Width           =   14835
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   735
         Index           =   3
         Left            =   360
         TabIndex        =   42
         Top             =   3480
         Width           =   14295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   735
         Index           =   1
         Left            =   360
         TabIndex        =   41
         Top             =   1560
         Width           =   14295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   735
         Index           =   2
         Left            =   360
         TabIndex        =   40
         Top             =   2520
         Width           =   14295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   735
         Index           =   0
         Left            =   360
         MaskColor       =   &H8000000F&
         TabIndex        =   3
         Top             =   660
         Width           =   14295
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   14940
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Options :"
         Height          =   555
         Left            =   180
         TabIndex        =   45
         ToolTipText     =   "Choose any one of the following"
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Timer Timer1 
      Left            =   180
      Top             =   60
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   1455
      Left            =   16320
      Top             =   60
      Width           =   4155
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   360
      X2              =   15300
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3060
      Picture         =   "exam.frx":304148
      Stretch         =   -1  'True
      ToolTipText     =   "Reads Question for you"
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Question"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   840
      TabIndex        =   50
      ToolTipText     =   "Question"
      Top             =   1980
      Width           =   2235
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Height          =   855
      Left            =   4380
      Shape           =   4  'Rounded Rectangle
      Top             =   10440
      Width           =   7815
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   8475
      Left            =   360
      Top             =   1860
      Width           =   14955
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   7875
      Left            =   15540
      Top             =   2460
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   3180
      TabIndex        =   47
      Top             =   60
      Width           =   11295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ONLINE EXAMINATION"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   3180
      TabIndex        =   46
      ToolTipText     =   "Test your Skils"
      Top             =   1020
      Width           =   11295
   End
End
Attribute VB_Name = "frmexam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim time As Single

'to keep track on the current question index(ID of report table)
Dim current As Single

'to chk if attempted
Dim gaveAns(29) As Boolean

'stores 30 randomly generated no.'s
Dim qnum(29) As String

'to calculate result i.e total marks
Dim RepliedAns(29) As String

'options
Dim opt(2 To 5) As Single


Public Sub getQuestion()
    Dim i As Single
    Dim rs As New ADODB.Recordset
    Set rs = entrypoint.conn.Execute("select * from " & Examtable & " where ID = " & qnum((current - 1)) & "")
    rs.MoveFirst
    
    'Load question on the screen
    question.Text = rs.Fields(1)
    
    GenerateRandomSequenceOfOptions
    
    Option1(0).Caption = rs.Fields(opt(2))
    Option1(1).Caption = rs.Fields(opt(3))
    Option1(2).Caption = rs.Fields(opt(4))
    Option1(3).Caption = rs.Fields(opt(5))
    
    
    'To avoid hanging of program
    opt(2) = 0
    opt(3) = 0
    opt(4) = 0
    opt(5) = 0
    
    If gaveAns(current - 1) = False Then
        'make all values false as question is not still attempted
        setVal
    Else
        'retain the option selected by user
        For i = 0 To 3
            If Option1(i).Caption = RepliedAns(current - 1) Then
                Option1(i).Value = True
                Exit For
            End If
        Next
    End If
    rs.Close
End Sub

Private Sub GenerateRandomSequenceOfOptions()
    Dim cnt As Single
    Dim n As Single
    cnt = 2
    
    Do While cnt <= 5
        n = CInt(Rnd * 5)
                
        If Not n < 2 Then
            If searchOpt(n) = False Then
                opt(cnt) = n
                cnt = cnt + 1
            End If
        End If
    Loop
    
End Sub

'search
Private Function searchOpt(n As Single) As Boolean
    Dim present As Boolean
    Dim i As Single
    present = False
    
    For i = LBound(opt) To UBound(opt)
        If opt(i) = n Then
            present = True
            GoTo e
        End If
    Next
e:
    searchOpt = present
End Function



Private Sub Command1_Click()    'submit Button & timer4
    Dim ask As VbMsgBoxResult
    If MsgBox("Do you want to SUBMIT ? ", vbYesNoCancel + vbQuestion, "finish") = vbYes Then
        calculate_result
        Me.Hide
        Timer4.Enabled = False
        Load frmResult
        frmResult.Show
        'Load SHORT_REPORT
        'SHORT_REPORT.Show
    End If
End Sub

Private Sub Command2_Click()    'Exit Button
    Unload Me
End Sub

Private Sub Command3_Click(Index As Integer)    'Change Question Button
    checkbutton
                                                                
    If Index = 0 Then                   'next question
        current = current + 1
        checkbutton
        
        
        Label1(0).Caption = current & " ) Question "
        
        getQuestion
    Else                                'previous question
        current = current - 1
        checkbutton
        
        Label1(0).Caption = current & " ) Question "
        
        getQuestion
    End If
    checkbutton
End Sub

Private Sub Command3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    Command3(Index).BackColor = vbRed
End Sub

Private Sub Command3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    Command3(Index).BackColor = vbButtonFace
End Sub

Private Sub Command4_Click(Index As Integer)    'All questions FRAME
       
    current = Index + 1 'i.e caption
    checkbutton
    Label1(0).Caption = current & " ) Question "
    
    getQuestion
End Sub

Private Sub Form_Load()
    Dim i As Single
    
    time = 1800
    frmlogin.Hide
    Me.Show

    For i = LBound(RepliedAns) To UBound(RepliedAns)
        gaveAns(i) = False
        RepliedAns(i) = vbNullString
       ' ActualAns(i) = vbNullString
    Next
    
    current = 1
    generate_questions
    checkbutton
    
    Command3(1).Enabled = False 'previous question button disabled
    Command4_Click (0) 'display 1st question by default
        
    Timer2.Interval = 1     'for no.of ans attempted
    Timer2.Enabled = True
        
    'exam caption
    Select Case entrypoint.Examtable
        Case "c"
            Label1(2).Caption = "C"
        Case "cpp"
            Label1(2).Caption = "C++"
        Case "csharp"
            Label1(2).Caption = "C#"
        Case Else
            Label1(2).Caption = entrypoint.Examtable
    End Select
        
    'Username
    Label2.Caption = " " & uname
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Load MDIadmin
        MDIadmin.Show
End Sub

Private Sub Image1_Click()  'sound
    Dim msg, sapi
    msg = question.Text
    Set sapi = CreateObject("sapi.spvoice")
    sapi.Speak msg
End Sub

Private Sub Option1_Click(Index As Integer)     'Option Clicked i.e ATTEMPTED
    Command4(current - 1).BackColor = vbGreen
    gaveAns(current - 1) = True
    RepliedAns(current - 1) = Option1(Index).Caption
End Sub

Private Sub setVal()    'make unattempted
        Option1(0).Value = False
        Option1(1).Value = False
        Option1(2).Value = False
        Option1(3).Value = False
End Sub

'============================================================================

Public Sub generate_questions()
    Dim rs As New ADODB.Recordset
    Dim n As Single
    Dim i As Single
    Dim sql As String
        
    For i = LBound(qnum) To UBound(qnum)
        qnum(i) = -1
    Next
    
    getnum  'get 30 random ID
    
   
    'get those questios from table
    For i = LBound(qnum) To UBound(qnum)
        Set rs = entrypoint.conn.Execute("select * from " & entrypoint.Examtable & " where ID = " & qnum(i) & "")
        rs.MoveFirst
    Next    'move to next row
    
    rs.Close
End Sub

'Get 30 random ID's
Private Sub getnum()
    Dim cnt As Single
    Dim n As Single
    Dim MaxRows As Single
    Dim rs As New ADODB.Recordset
    
    
    Set rs = conn.Execute("select max(ID) from " & Examtable)
    rs.MoveFirst
    MaxRows = rs.Fields(0)
    cnt = 0
    
    Do While cnt <= UBound(qnum)
        n = CInt(Rnd * MaxRows)  'total 100 questions in table
        
        If search(n) = False And n <> 0 Then
            qnum(cnt) = n
            cnt = cnt + 1
        End If
    Loop
    rs.Close
End Sub

'search
Private Function search(n As Single) As Boolean
    Dim present As Boolean
    Dim i As Single
    present = False
    
    For i = LBound(qnum) To UBound(qnum)
        If qnum(i) = n Then
            present = True
            GoTo e
        End If
    Next
e:
    search = present
End Function

Private Sub checkbutton()   'Enable and disable of NEXT and PREVIOUS button
    
    If current = 30 Then        'Next question
        Command3(0).Enabled = False
    Else
        Command3(0).Enabled = True
    End If
        
    If current = 1 Then        'prev question
        Command3(1).Enabled = False
    Else
        Command3(1).Enabled = True
    End If
    
End Sub



Private Sub calculate_result()      'Result Calculation
    Dim i As Single
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
    score = 0
       
    For i = LBound(qnum) To UBound(qnum)
        sql = "select answer from  "
        sql = sql & Examtable & " where ID = " & qnum(i)
        rs.Open sql, conn, adOpenDynamic, adLockOptimistic
        rs.MoveFirst
       
        If gaveAns(i) = True Then
            If RepliedAns(i) = rs.Fields(0) Then
                score = score + 1
            End If
        End If
        
        rs.Close
    Next
    
    sql = "update report set score = " & score & " where rid = '" & uname & "'"
    conn.Execute sql
    
   
    'conn.Execute "commit"
    
    'rs.Close
End Sub

Private Sub Timer2_Timer()  'Update no.of questions attempted
    Dim i As Single
    Dim cnt As Single
    
    cnt = 0
    For i = 0 To 29
        If Command4(i).BackColor = vbGreen Then
            cnt = cnt + 1
        End If
    Next
    Label4.Caption = "  " & cnt & " / 30 "
    
    
End Sub

Private Sub Timer4_Timer()  'time up
    time = time - 1
    
    If time = 0 Then
        MsgBox "Do you want to SUBMIT ? ", vbOKOnly + vbQuestion, "finish"
        calculate_result
        Me.Hide
        Load frmResult
        frmResult.Show
        'Load SHORT_REPORT
        'SHORT_REPORT.Show
    End If
End Sub
