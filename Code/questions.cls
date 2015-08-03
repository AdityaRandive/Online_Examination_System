VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "questions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



'Constructor
Public Sub generate_questions()

    Dim n As Single
    Dim i As Single
        
    For i = LBound(qnum) To UBound(qnum)
        qnum(i) = -1
    Next
    
    getnum  'get 30 random ID
    
    'get those questios from table
    For i = LBound(qnum) To UBound(qnum)
        Set rs = entrypoint.conn.Execute("select * from " & entrypoint.Examtable & " where ID = " & qnum(i) & ";")
        
        rs.MoveFirst
        'fetch data of cols of current row
                
        qn(i) = rs.Fields(1)
        opt1(i) = rs.Fields(2)
        opt2(i) = rs.Fields(3)
        opt3(i) = rs.Fields(4)
        ans(i) = rs.Fields(5)
            
        'move to next row
    Next
    
    rs.Close
End Sub


'Get 30 random ID's
Private Sub getnum()
    Dim cnt As Single
    Dim n As Single
    
    cnt = 0
    
    Do While cnt <= UBound(qnum)
        n = CInt(Rnd * 100)  'total 100 questions in table
        
        If search(n) = False Then
            qnum(cnt) = n
            cnt = cnt + 1
        End If
    Loop
            
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

Public Function get_qn(n As Integer)
    get_n = qn(Index)
End Function

Public Function get_opt1(n As Integer)
    get_opt1 = opt1(Index)
End Function

Public Function get_opt2(n As Integer)
    get_opt2 = opt2(Index)
End Function

Public Function get_opt3(n As Integer)
    get_opt3 = opt3(Index)
End Function

Public Function get_ans(n As Integer)
    get_ans = ans(Index)
End Function


