VERSION 5.00
Begin VB.MDIForm MDIadmin 
   BackColor       =   &H8000000C&
   Caption         =   "Administrator"
   ClientHeight    =   8265
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11115
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "MDIadmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    
    If entrypoint.LoginType = 3 Then
        frmAdmin.Show
        entrypoint.LoginType = 0
    ElseIf entrypoint.CreateAccount = True Then
        frmcreate.Show
        entrypoint.CreateAccount = False
    Else    'Student Login type = 0
        Load frmStudent
        frmStudent.Show
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    frmlogin.Show
    conn.Execute "delete from report where rid = '" & uname & "'"
End Sub
